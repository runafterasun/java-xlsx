package ru.objectsfill.dto;

import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.experimental.FieldNameConstants;

@Data
@NoArgsConstructor
@FieldNameConstants
public class ExcelImport {

    private String account;
    private String offerDate;
    private String currency;
    private String product;
    private String rateModNumber;

}
