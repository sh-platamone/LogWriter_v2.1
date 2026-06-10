package org.example.service;

import org.example.model.LocalAccount;

import javax.crypto.SecretKeyFactory;
import javax.crypto.spec.PBEKeySpec;
import java.io.IOException;
import java.io.Reader;
import java.io.Writer;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.security.MessageDigest;
import java.security.NoSuchAlgorithmException;
import java.security.SecureRandom;
import java.security.spec.InvalidKeySpecException;
import java.time.Instant;
import java.util.Arrays;
import java.util.Base64;
import java.util.Locale;
import java.util.Properties;
import java.util.regex.Pattern;

import static java.nio.file.StandardOpenOption.CREATE;
import static java.nio.file.StandardOpenOption.TRUNCATE_EXISTING;
import static java.nio.file.StandardOpenOption.WRITE;
import static java.nio.file.StandardCopyOption.REPLACE_EXISTING;

public class LocalAccountService {

    private static final Pattern USERNAME_PATTERN = Pattern.compile("[A-Za-z0-9._@-]{3,64}");
    private static final int PASSWORD_MIN_LENGTH = 8;
    private static final int SALT_BYTES = 16;
    private static final int HASH_ITERATIONS = 65_536;
    private static final int HASH_BITS = 256;

    private final SecureRandom secureRandom = new SecureRandom();
    private final Path storageDirectory;
    private final Path accountsFile;
    private final Path usersDirectory;

    public LocalAccountService() {
        this(ApplicationDataDirectoryResolver.resolveDataDirectory());
    }

    LocalAccountService(Path storageDirectory) {
        this.storageDirectory = storageDirectory;
        this.accountsFile = storageDirectory.resolve("accounts.properties");
        this.usersDirectory = storageDirectory.resolve("users");
        migrateLegacyStorageIfNeeded();
    }

    public synchronized LocalAccount createAccount(String username, String password) throws IOException {
        String sanitizedUsername = sanitizeUsername(username);
        validatePassword(password);

        Properties properties = loadProperties();
        String canonicalUsername = canonicalizeUsername(sanitizedUsername);
        String propertyPrefix = buildPropertyPrefix(canonicalUsername);
        if (properties.containsKey(propertyPrefix + "hash")) {
            throw new IllegalArgumentException("An account with that username already exists.");
        }

        byte[] salt = new byte[SALT_BYTES];
        secureRandom.nextBytes(salt);
        byte[] passwordHash = hashPassword(password.toCharArray(), salt, HASH_ITERATIONS);

        properties.setProperty(propertyPrefix + "username", sanitizedUsername);
        properties.setProperty(propertyPrefix + "salt", Base64.getEncoder().encodeToString(salt));
        properties.setProperty(propertyPrefix + "hash", Base64.getEncoder().encodeToString(passwordHash));
        properties.setProperty(propertyPrefix + "iterations", Integer.toString(HASH_ITERATIONS));
        properties.setProperty(propertyPrefix + "createdAt", Instant.now().toString());
        saveProperties(properties);

        return createLocalAccount(sanitizedUsername, canonicalUsername);
    }

    public synchronized LocalAccount authenticate(String username, String password) throws IOException {
        String sanitizedUsername = sanitizeUsername(username);
        validatePassword(password);

        Properties properties = loadProperties();
        String canonicalUsername = canonicalizeUsername(sanitizedUsername);
        String propertyPrefix = buildPropertyPrefix(canonicalUsername);
        String storedHash = properties.getProperty(propertyPrefix + "hash");
        String storedSalt = properties.getProperty(propertyPrefix + "salt");
        String storedIterations = properties.getProperty(propertyPrefix + "iterations");

        if (storedHash == null || storedSalt == null || storedIterations == null) {
            throw new IllegalArgumentException("Invalid username or password.");
        }

        byte[] expectedHash = Base64.getDecoder().decode(storedHash);
        byte[] salt = Base64.getDecoder().decode(storedSalt);
        int iterations = parseIterations(storedIterations);
        byte[] actualHash = hashPassword(password.toCharArray(), salt, iterations);
        if (!MessageDigest.isEqual(expectedHash, actualHash)) {
            throw new IllegalArgumentException("Invalid username or password.");
        }

        String displayUsername = properties.getProperty(propertyPrefix + "username", sanitizedUsername);
        return createLocalAccount(displayUsername, canonicalUsername);
    }

    private LocalAccount createLocalAccount(String displayUsername, String canonicalUsername) throws IOException {
        Path userDataDirectory = usersDirectory.resolve(canonicalUsername);
        Files.createDirectories(userDataDirectory);
        return new LocalAccount(displayUsername, userDataDirectory);
    }

    private Properties loadProperties() throws IOException {
        Properties properties = new Properties();
        if (!Files.exists(accountsFile)) {
            return properties;
        }

        try (Reader reader = Files.newBufferedReader(accountsFile, StandardCharsets.UTF_8)) {
            properties.load(reader);
        }
        return properties;
    }

    private void saveProperties(Properties properties) throws IOException {
        Files.createDirectories(storageDirectory);
        Files.createDirectories(usersDirectory);
        try (Writer writer = Files.newBufferedWriter(accountsFile, StandardCharsets.UTF_8, CREATE, TRUNCATE_EXISTING, WRITE)) {
            properties.store(writer, "LogWriter local accounts");
        }
    }

    private String sanitizeUsername(String username) {
        String sanitizedUsername = username == null ? "" : username.trim();
        if (!USERNAME_PATTERN.matcher(sanitizedUsername).matches()) {
            throw new IllegalArgumentException("Usernames must be 3 to 64 characters using letters, numbers, '.', '_', '-', or '@'.");
        }
        return sanitizedUsername;
    }

    private void validatePassword(String password) {
        if (password == null || password.length() < PASSWORD_MIN_LENGTH) {
            throw new IllegalArgumentException("Passwords must be at least %d characters long.".formatted(PASSWORD_MIN_LENGTH));
        }
    }

    private int parseIterations(String storedIterations) {
        try {
            return Integer.parseInt(storedIterations);
        } catch (NumberFormatException exception) {
            throw new IllegalStateException("The stored account data is invalid.");
        }
    }

    private byte[] hashPassword(char[] password, byte[] salt, int iterations) {
        PBEKeySpec keySpec = new PBEKeySpec(password, salt, iterations, HASH_BITS);
        Arrays.fill(password, '\0');
        try {
            return SecretKeyFactory.getInstance("PBKDF2WithHmacSHA256")
                    .generateSecret(keySpec)
                    .getEncoded();
        } catch (NoSuchAlgorithmException | InvalidKeySpecException exception) {
            throw new IllegalStateException("Unable to secure local account passwords on this device.", exception);
        } finally {
            keySpec.clearPassword();
        }
    }

    private String canonicalizeUsername(String username) {
        return username.toLowerCase(Locale.ROOT);
    }

    private String buildPropertyPrefix(String canonicalUsername) {
        return "account.%s.".formatted(canonicalUsername);
    }

    private void migrateLegacyStorageIfNeeded() {
        Path legacyStorageDirectory = resolveLegacyStorageDirectory();
        Path legacyAccountsFile = legacyStorageDirectory.resolve("accounts.properties");
        Path legacyUsersDirectory = legacyStorageDirectory.resolve("users");

        try {
            if (!Files.exists(accountsFile) && Files.exists(legacyAccountsFile)) {
                Files.createDirectories(storageDirectory);
                Files.move(legacyAccountsFile, accountsFile, REPLACE_EXISTING);
            }

            if (!Files.exists(usersDirectory) && Files.isDirectory(legacyUsersDirectory)) {
                Files.createDirectories(storageDirectory);
                Files.move(legacyUsersDirectory, usersDirectory);
            }
        } catch (IOException exception) {
            throw new IllegalStateException("Unable to migrate existing local account data to the app data folder.", exception);
        }
    }

    private static Path resolveLegacyStorageDirectory() {
        String appData = System.getenv("APPDATA");
        if (appData != null && !appData.isBlank()) {
            return Paths.get(appData, "LogWriter");
        }

        return Paths.get(System.getProperty("user.home"), ".logwriter");
    }
}
