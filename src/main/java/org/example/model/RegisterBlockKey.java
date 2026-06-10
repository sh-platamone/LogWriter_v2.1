package org.example.model;

public record RegisterBlockKey(int rowIndex, int firstColumnIndex) {

    public int secondColumnIndex() {
        return firstColumnIndex + 1;
    }
}
