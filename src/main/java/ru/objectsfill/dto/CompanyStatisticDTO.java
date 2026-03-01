package ru.objectsfill.dto;

import lombok.Data;

import java.util.List;


@Data
public class CompanyStatisticDTO {

    private String companyId;
    private String companyName;
    private String commercialStatus;
    private List<String> houseStatisticList;
    private Long workerCount = 0L;
    private Long supervisorCount = 0L;
    private Long masterCount = 0L;

}
