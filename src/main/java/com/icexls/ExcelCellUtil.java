package com.icexls;

public class ExcelCellUtil {

    static boolean isNumber(String str) {
        if (str == null || "".equals(str.trim())) {
            return false;
        }
        if (str.charAt(0) == '-') {
            str = str.substring(1);
        }
        int point = 0;
        for (int i = 0; i < str.length(); i++) {
            char ch = str.charAt(i);
            if (ch == '.') {
                point++;
            } else if (ch < '0' || ch > '9') {
                return false;
            }
        }
        if (point > 1) {
            return false;
        }
        return true;
    }
}
