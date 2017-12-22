package com.icexls;

import java.util.List;
/**
 * List<String[]>与String[][]的转换
 * 
 * @author iceWater
 * @date 2017-12-15
 * @version 2.0
 */
public class IceExcelConvert {
    public static String[][] convert(List<String[]> list) {
        if (list == null || list.size() == 0) {
            return new String[1][1];
        } else {
            int rowLength = list.size();
            String[][] data = new String[rowLength][];
            for (int i = 0; i < list.size(); i++) {
                String[] row = list.get(i);
                data[i] = row;
            }
            return data;
        }
    }
}
