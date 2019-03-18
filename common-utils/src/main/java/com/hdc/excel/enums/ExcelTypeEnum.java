package com.hdc.excel.enums;

/**
 * ExcelTypeEnum
 * @author hdc
 * @date 2019/03/14
 */
public enum ExcelTypeEnum {
    XLS(1, "xls"),
    XLSX(2, "xlsx");

    public Integer code;
    public String value;

    ExcelTypeEnum(Integer code, String value) {
        this.code = code;
        this.value = value;
    }

    public static ExcelTypeEnum getDescByValue(String value){
        for (ExcelTypeEnum ete : ExcelTypeEnum.values()) {
            if (ete.value.equals(value)) {
                return ete;
            }
        }
        return null;
    }
}
