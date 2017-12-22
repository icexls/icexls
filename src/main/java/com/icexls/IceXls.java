package com.icexls;
/**
 * 导出Excel的工具类
 * 
 * @author iceWater
 * @date 2017-12-22
 * @version 2.0
 */
public class IceXls extends IceExcel {

    public IceXls(String excelFileName) {
        super(excelFileName);
    }

    public IceXls(String excelFileName, String sheetName) {
        super(excelFileName, sheetName);
    }
}
