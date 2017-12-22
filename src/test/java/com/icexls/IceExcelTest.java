package com.icexls;

import org.junit.Test;

import com.icexls.IceExcel;
import com.icexls.IceExcelConfig;
import com.icexls.NumberType;
import com.icexls.ParserType;

public class IceExcelTest {
    @Test
    public void export() {
        String[][] data = { { "aaa", "167" }, { "278", "bbb2" }, { "aaa3", "120.36" }, { "aaa4", "bbb4" } };
        IceExcel iceExcel = new IceExcel("D:/xls-test.xls");
        // IceExcel("C:/Users/Administrator/Desktop/xls-test.xls","test-data");
        // IceExcelConfig.setSheet(iceExcel,"hello");
        IceExcelConfig.setNumberType(iceExcel, NumberType.STRING);
        IceExcelConfig.setParserType(iceExcel, ParserType.POI);
        iceExcel.setData(data);
    }

    // @Test
    public void importx() {
        IceExcel iceExcel = new IceExcel("D:/xls-test.xls");
        IceExcelConfig.setParserType(iceExcel, ParserType.POI);
        String[][] data = iceExcel.getData();
        for (int i = 0; i < data.length; i++) {
            for (int j = 0; j < data[i].length; j++) {
                System.out.print(data[i][j] + "\t  ");
            }
            System.out.println();
        }
    }
}
