package com.icexls;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

public class PoiExcelParser extends AbstractExcelParser implements ExcelParser {

    // @Override
    public String[][] getData() {
        String sheetNameCurrent = this.getSheet();
        BufferedInputStream bufferedInputStream = null;
        try {
            bufferedInputStream = new BufferedInputStream(new FileInputStream(this.getExcelFileName()));
        } catch (FileNotFoundException e1) {
            e1.printStackTrace();
        }
        HSSFWorkbook hssfWorkbook = null;
        try {
            POIFSFileSystem poifsFileSystem = new POIFSFileSystem(bufferedInputStream);
            hssfWorkbook = new HSSFWorkbook(poifsFileSystem);
        } catch (IOException e) {
            e.printStackTrace();
        }
        HSSFSheet sheetAtCurrent = null;
        if (sheetNameCurrent != null) {
            for (int sheetIndex = 0; sheetIndex < hssfWorkbook.getNumberOfSheets(); sheetIndex++) {
                HSSFSheet sheetAt = hssfWorkbook.getSheetAt(sheetIndex);
                if (sheetAt == null) {
                    continue;
                }
                String sheetName = sheetAt.getSheetName();
                if (sheetNameCurrent.equals(sheetName)) {
                    sheetAtCurrent = sheetAt;
                    break;
                }
            }
        } else {
            sheetAtCurrent = hssfWorkbook.getSheetAt(0);
        }
        int lastRowNum = sheetAtCurrent.getLastRowNum();
        String[][] result = new String[lastRowNum][];
        for (int i = 0; i < lastRowNum; i++) {
            HSSFRow row = sheetAtCurrent.getRow(i);
            if (row == null) {
                continue;
            }
            short lastCellNum = row.getLastCellNum();
            result[i] = new String[lastCellNum];
            for (int j = 0; j < lastCellNum; j++) {
                HSSFCell cell = row.getCell(j);
                // cell.
                int cellType = cell.getCellType();
                if (cellType == HSSFCell.CELL_TYPE_STRING) {
                    String stringCellValue = cell.getStringCellValue();
                    result[i][j] = stringCellValue;
                } else if (cellType == HSSFCell.CELL_TYPE_NUMERIC) {
                    if (HSSFDateUtil.isCellDateFormatted(cell)) {
                        Date dateCellValue = cell.getDateCellValue();
                        if (dateCellValue == null) {
                            result[i][j] = "";
                        } else {
                            SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd");
                            result[i][j] = simpleDateFormat.format(dateCellValue);
                        }
                    } else {
                        double numericCellValue = cell.getNumericCellValue();
                        result[i][j] = numericCellValue + "";
                    }
                } else if (cellType == HSSFCell.CELL_TYPE_BOOLEAN) {
                    boolean booleanCellValue = cell.getBooleanCellValue();
                    result[i][j] = booleanCellValue ? "Y" : "N";
                } else if (cellType == HSSFCell.CELL_TYPE_ERROR) {
                    result[i][j] = "";
                } else if (cellType == HSSFCell.CELL_TYPE_FORMULA) {
                    String stringCellValue = cell.getStringCellValue();
                    double numericCellValue = cell.getNumericCellValue();
                    if (!"".equals(stringCellValue)) {
                        result[i][j] = stringCellValue;
                    } else {
                        result[i][j] = numericCellValue + "";
                    }
                } else if (cellType == HSSFCell.CELL_TYPE_BLANK) {
                    result[i][j] = "";
                }
            }
        }
        if (bufferedInputStream != null) {
            try {
                bufferedInputStream.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return result;
    }

    // @Override
    public void setData(String[][] data) {
        HSSFWorkbook hssfWorkbook = new HSSFWorkbook();
        String sheetname = this.getSheet();
        HSSFSheet createSheet = hssfWorkbook.createSheet(sheetname);
        for (int i = 0; i < data.length; i++) {
            HSSFRow createRow = createSheet.createRow(i);
            for (int j = 0; j < data[i].length; j++) {
                HSSFCell createCell = createRow.createCell(j);
                String numberType = this.getNumberType();
                String cell = data[i][j];
                if ("NUMBER".equalsIgnoreCase(numberType) && ExcelCellUtil.isNumber(cell)) {
                    boolean idDouble = cell.indexOf(".") >= 0;
                    if (idDouble) {
                        double dou = Double.parseDouble(cell);
                        createCell.setCellValue(dou);
                    } else {
                        int num = Integer.parseInt(cell);
                        createCell.setCellValue(num);
                    }
                } else {
                    String value = "" + data[i][j];
                    createCell.setCellValue(value);
                }
            }
        }
        String xlsFileName = this.getExcelFileName();
        FileOutputStream fileOutputStream = null;
        File xlsFile = new File(xlsFileName);
        File parentFile = xlsFile.getParentFile();
        if(!parentFile.exists()){
            parentFile.mkdirs();
        }
        try {
            fileOutputStream = new FileOutputStream(xlsFileName);
            hssfWorkbook.write(fileOutputStream);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (fileOutputStream != null) {
                try {
                    fileOutputStream.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
        try {
            if (hssfWorkbook != null) {
                hssfWorkbook.close();
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
