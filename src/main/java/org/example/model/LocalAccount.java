package org.example.model;

import java.nio.file.Path;

public record LocalAccount(
        String username,
        Path dataDirectory
) {
}
