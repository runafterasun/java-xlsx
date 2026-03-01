package ru.objectsfill.enums;

public enum LoopOperationType {

    FOR("for."),
    ROW("row");

    private final String type;

    LoopOperationType(String type) {
        this.type = type;
    }

    public String getType() {
        return type;
    }
}
