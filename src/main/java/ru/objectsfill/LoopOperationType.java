package ru.objectsfill;


import lombok.AllArgsConstructor;
import lombok.Getter;

@Getter
@AllArgsConstructor
public enum LoopOperationType {

    FOR("for"),
    ROW("row");

    private final String type;
}
