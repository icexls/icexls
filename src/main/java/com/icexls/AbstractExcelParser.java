package com.icexls;

public abstract class AbstractExcelParser {
    private String excelFileName;
    private String sheet;
    private String numberType;

    protected void setExcelFileName(String excelFileName) {
        this.excelFileName = excelFileName;
    }

    protected void setSheet(String sheet) {
        this.sheet = sheet;
    }

    public String getExcelFileName() {
        return excelFileName;
    }

    public String getSheet() {
        return sheet;
    }

    public void setNumberType(NumberType numberType) {
        this.numberType = numberType + "";
    }

    public String getNumberType() {
        return numberType;
    }
}
