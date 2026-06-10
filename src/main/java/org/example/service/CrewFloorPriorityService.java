package org.example.service;

import org.example.model.LocalAccount;

import java.io.IOException;
import java.io.Reader;
import java.io.Writer;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.HashMap;
import java.util.Locale;
import java.util.Map;
import java.util.Properties;

import static java.nio.file.StandardOpenOption.CREATE;
import static java.nio.file.StandardOpenOption.TRUNCATE_EXISTING;
import static java.nio.file.StandardOpenOption.WRITE;

public class CrewFloorPriorityService {

    private static final String FILE_NAME = "crew-floor-priorities.properties";

    public synchronized Map<String, Integer> loadFloorPriorities(LocalAccount account) throws IOException {
        Properties properties = loadProperties(resolveStoragePath(account));
        Map<String, Integer> floorPriorities = new HashMap<>();
        for (String propertyName : properties.stringPropertyNames()) {
            try {
                floorPriorities.put(propertyName, Integer.parseInt(properties.getProperty(propertyName)));
            } catch (NumberFormatException ignored) {
            }
        }
        return floorPriorities;
    }

    public synchronized void saveFloorPriority(LocalAccount account, String crewMemberName, int floorPriority) throws IOException {
        Path storagePath = resolveStoragePath(account);
        Properties properties = loadProperties(storagePath);
        String propertyName = normalizeName(crewMemberName);

        if (floorPriority <= 0) {
            properties.remove(propertyName);
        } else {
            properties.setProperty(propertyName, Integer.toString(floorPriority));
        }

        Files.createDirectories(storagePath.getParent());
        try (Writer writer = Files.newBufferedWriter(storagePath, StandardCharsets.UTF_8, CREATE, TRUNCATE_EXISTING, WRITE)) {
            properties.store(writer, "Crew floor priority settings");
        }
    }

    private Properties loadProperties(Path storagePath) throws IOException {
        Properties properties = new Properties();
        if (!Files.exists(storagePath)) {
            return properties;
        }

        try (Reader reader = Files.newBufferedReader(storagePath, StandardCharsets.UTF_8)) {
            properties.load(reader);
        }
        return properties;
    }

    private Path resolveStoragePath(LocalAccount account) {
        if (account == null) {
            throw new IllegalArgumentException("A logged-in account is required to load crew settings.");
        }
        return account.dataDirectory().resolve(FILE_NAME);
    }

    public String normalizeName(String crewMemberName) {
        return crewMemberName == null ? "" : crewMemberName.trim().toLowerCase(Locale.ROOT);
    }
}
