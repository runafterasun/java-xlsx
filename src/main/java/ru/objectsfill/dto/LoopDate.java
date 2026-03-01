package ru.objectsfill.dto;

import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@NoArgsConstructor
public class LoopDate {

    private String destination;
    private String origin;
    private String code;
    private String route;
    private String date;
    private String rate;
    private String change;
    private String billing;


    private Integer row;


}
