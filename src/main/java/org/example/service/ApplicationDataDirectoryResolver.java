package org.example.service;

import java.net.URISyntaxException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

public final class ApplicationDataDirectoryResolver {

    private ApplicationDataDirectoryResolver() {
    }

    public static Path resolveDataDirectory() {
        return resolveApplicationRoot().resolve("data");
    }

    private static Path resolveApplicationRoot() {
        try {
            Path codeSourcePath = Paths.get(
                    ApplicationDataDirectoryResolver.class.getProtectionDomain().getCodeSource().getLocation().toURI()
            ).toAbsolutePath().normalize();

            if (Files.isRegularFile(codeSourcePath)) {
                Path parentDirectory = codeSourcePath.getParent();
                if (parentDirectory != null && parentDirectory.getFileName() != null) {
                    String parentDirectoryName = parentDirectory.getFileName().toString();
                    if ("app".equalsIgnoreCase(parentDirectoryName) || "target".equalsIgnoreCase(parentDirectoryName)) {
                        Path installRoot = parentDirectory.getParent();
                        if (installRoot != null) {
                            return installRoot;
                        }
                    }
                }
                return parentDirectory;
            }

            Path fileName = codeSourcePath.getFileName();
            Path parentDirectory = codeSourcePath.getParent();
            if (fileName != null
                    && "classes".equalsIgnoreCase(fileName.toString())
                    && parentDirectory != null
                    && parentDirectory.getFileName() != null
                    && "target".equalsIgnoreCase(parentDirectory.getFileName().toString())) {
                Path projectRoot = parentDirectory.getParent();
                if (projectRoot != null) {
                    return projectRoot;
                }
            }

            return codeSourcePath;
        } catch (URISyntaxException exception) {
            throw new IllegalStateException("Unable to determine the LogWriter installation folder.", exception);
        }
    }
}
