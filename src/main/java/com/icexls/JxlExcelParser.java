package com.icexls;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class JxlExcelParser extends AbstractExcelParser implements ExcelParser {

    @Override
    public String[][] getData() {
        File file = new File(this.getExcelFileName());
        Workbook workbook = null;
        try {
            workbook = Workbook.getWorkbook(file);
        } catch (BiffException e) {
            if ("Unable to recognize OLE stream".equals(e.getMessage())) {
                throw new RuntimeException("jxl暂不支持Excel 2007,请使用POI实现");
            } else {
                e.printStackTrace();
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        String sheetName = this.getSheet();
        String[] sheetNames = workbook.getSheetNames();
        Sheet sheet = null;
        if (sheetName != null) {
            for (String sheetNameInArray : sheetNames) {
                if (sheetName.equals(sheetNameInArray)) {
                    sheet = workbook.getSheet(sheetName);
                }
            }
        } else {
            sheet = workbook.getSheet(0);// Excel本身控制不会越界
        }
        int rowLength = sheet.getRows();
        String[][] result = new String[rowLength][];
        for (int i = 0; i < rowLength; i++) {
            Cell[] rows = sheet.getRow(i);
            int coumnLength = rows.length;
            result[i] = new String[coumnLength];
            for (int j = 0; j < coumnLength; j++) {
                Cell cell = rows[j];
                String contents = cell.getContents();
                result[i][j] = contents;
            }
        }
        if (workbook != null) {
            workbook.close();
        }
        return result;
    }

    @Override
    public void setData(String[][] data) {
        String xlsFileName = this.getExcelFileName();
        File file = new File(xlsFileName);
        String sheetName = this.getSheet();
        String numberType = this.getNumberType();
        WritableWorkbook createWorkbook = null;
        try {
            File xlsFile = new File(xlsFileName);
            File parentFile = xlsFile.getParentFile();
            if (!parentFile.exists()) {
                parentFile.mkdirs();
            }
            try {
                createWorkbook = Workbook.createWorkbook(file);
            } catch (FileNotFoundException e) {
                if (file.isFile() && file.exists()) {
                    throw new RuntimeException("Excel: " + file.getAbsolutePath().replace("\\", "/") + " 已打开");
                } else {
                    throw new RuntimeException(e);
                }
            }
            WritableSheet createSheet = createWorkbook.createSheet(sheetName, 0);
            for (int i = 0; i < data.length; i++) {
                for (int j = 0; j < data[i].length; j++) {
                    String cell = data[i][j];
                    if ("NUMBER".equalsIgnoreCase(numberType) && ExcelCellUtil.isNumber(cell)) {
                        boolean idDouble = cell.indexOf(".") >= 0;
                        if (idDouble) {
                            double dou = Double.parseDouble(cell);
                            createSheet.addCell(new jxl.write.Number(j, i, dou));
                        } else {
                            int num = Integer.parseInt(cell);
                            createSheet.addCell(new jxl.write.Number(j, i, num));
                        }
                    } else {
                        createSheet.addCell(new Label(j, i, cell));
                    }
                }
            }
            createWorkbook.write();
        } catch (RowsExceededException e) {
            e.printStackTrace();
        } catch (WriteException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (createWorkbook != null) {
                try {
                    createWorkbook.close();
                } catch (WriteException e) {
                    e.printStackTrace();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }
}
