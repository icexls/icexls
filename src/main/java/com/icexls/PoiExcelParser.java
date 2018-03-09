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
import org.apache.poi.poifs.filesystem.OfficeXmlFileException;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class PoiExcelParser extends AbstractExcelParser implements ExcelParser {

    @Override
    public String[][] getData() {
        String sheetNameCurrent = this.getSheet();
        BufferedInputStream bufferedInputStream = null;
        try {
            bufferedInputStream = new BufferedInputStream(new FileInputStream(this.getExcelFileName()));
        } catch (FileNotFoundException e1) {
            e1.printStackTrace();
        }
        Workbook hssfWorkbook = null;
        try {
            POIFSFileSystem poifsFileSystem = null;
            try {
                poifsFileSystem = new POIFSFileSystem(bufferedInputStream);
                hssfWorkbook = new HSSFWorkbook(poifsFileSystem);
            } catch (OfficeXmlFileException e) {
                try {
                    hssfWorkbook = new XSSFWorkbook(new FileInputStream(this.getExcelFileName()));
                } catch (NoClassDefFoundError e1) {
                    if ("org/apache/poi/xssf/usermodel/XSSFWorkbook".equals(e1.getMessage().trim())) {
                        throw new RuntimeException(
                                "没有引入poi-ooxml-x.x.jar,你可以从下面的地址下载:http://central.maven.org/maven2/org/apache/poi/poi-ooxml/3.17/poi-ooxml-3.17.jar");
                    } else if ("org/apache/xmlbeans/XmlObject".equals(e1.getMessage().trim())) {
                        throw new RuntimeException(
                                "没有引入xmlbeans-x.x.x.jar,你可以从下面的地址下载:http://central.maven.org/maven2/org/apache/xmlbeans/xmlbeans/2.6.0/xmlbeans-2.6.0.jar");
                    } else if ("org/apache/commons/collections4/ListValuedMap".equals(e1.getMessage().trim())) {
                        throw new RuntimeException(
                                "没有引入commons-collections4-x.x.jar,你可以从下面的地址下载:http://central.maven.org/maven2/org/apache/commons/commons-collections4/4.1/commons-collections4-4.1.jar");
                    } else if ("org/openxmlformats/schemas/drawingml/x2006/main/ThemeDocument"
                            .equals(e1.getMessage().trim())) {
                        throw new RuntimeException(
                                "没有引入ooxml-schemas-x.x.jar,你可以从下面的地址下载:http://central.maven.org/maven2/org/apache/poi/ooxml-schemas/1.3/ooxml-schemas-1.3.jar");
                    } else {
                        e1.printStackTrace();
                    }
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        Sheet sheetAtCurrent = null;
        if (sheetNameCurrent != null) {
            for (int sheetIndex = 0; sheetIndex < hssfWorkbook.getNumberOfSheets(); sheetIndex++) {
                Sheet sheetAt = hssfWorkbook.getSheetAt(sheetIndex);
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
        String[][] result = new String[lastRowNum + 1][];
        for (int i = 0; i <= lastRowNum; i++) {
            Row row = sheetAtCurrent.getRow(i);
            if (row == null) {
                continue;
            }
            short lastCellNum = row.getLastCellNum();
            result[i] = new String[lastCellNum];
            for (int j = 0; j < lastCellNum; j++) {
                Cell cell = row.getCell(j);
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

    @Override
    public void setData(String[][] data) {
        String xlsFileName = this.getExcelFileName();
        if (xlsFileName != null && xlsFileName.endsWith(".xlsx")) {
            excel2007Write(data, xlsFileName);
            return;
        }
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

        FileOutputStream fileOutputStream = null;
        File xlsFile = new File(xlsFileName);
        File parentFile = xlsFile.getParentFile();
        if (!parentFile.exists()) {
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

    private void excel2007Write(String[][] data, String xlsFileName) {
        XSSFWorkbook xssFWorkbook = null;
        try {
            xssFWorkbook = new XSSFWorkbook();
        } catch (NoClassDefFoundError e) {
            if ("org/apache/poi/xssf/usermodel/XSSFWorkbook".equals(e.getMessage().trim())) {
                throw new RuntimeException(
                        "没有引入poi-ooxml-x.x.jar,你可以从下面的地址下载:http://central.maven.org/maven2/org/apache/poi/poi-ooxml/3.17/poi-ooxml-3.17.jar");
            } else if ("org/apache/xmlbeans/XmlObject".equals(e.getMessage().trim())) {
                throw new RuntimeException(
                        "没有引入xmlbeans-x.x.x.jar,你可以从下面的地址下载:http://central.maven.org/maven2/org/apache/xmlbeans/xmlbeans/2.6.0/xmlbeans-2.6.0.jar");
            } else if ("org/apache/commons/collections4/ListValuedMap".equals(e.getMessage().trim())) {
                throw new RuntimeException(
                        "没有引入commons-collections4-x.x.jar,你可以从下面的地址下载:http://central.maven.org/maven2/org/apache/commons/commons-collections4/4.1/commons-collections4-4.1.jar");
            } else if ("org/openxmlformats/schemas/spreadsheetml/x2006/main/CTWorkbook$Factory"
                    .equals(e.getMessage().trim())) {
                throw new RuntimeException(
                        "没有引入poi-ooxml-schemas-x.x.jar,你可以从下面的地址下载:http://central.maven.org/maven2/org/apache/poi/poi-ooxml-schemas/3.17/poi-ooxml-schemas-3.17.jar");
            } else if ("org/openxmlformats/schemas/drawingml/x2006/main/ThemeDocument".equals(e.getMessage().trim())) {
                throw new RuntimeException(
                        "没有引入ooxml-schemas-x.x.jar,你可以从下面的地址下载:http://central.maven.org/maven2/org/apache/poi/ooxml-schemas/1.3/ooxml-schemas-1.3.jar");
            } else {
                e.printStackTrace();
            }
        }
        String sheetname = this.getSheet();
        XSSFSheet createSheet = xssFWorkbook.createSheet(sheetname);
        String numberType = this.getNumberType();
        for (int i = 0; i < data.length; i++) {
            XSSFRow createRow = createSheet.createRow(i);
            for (int j = 0; j < data[i].length; j++) {
                XSSFCell createCell = createRow.createCell(j);
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
        FileOutputStream fos = null;
        try {
            fos = new FileOutputStream(xlsFileName);
            xssFWorkbook.write(fos);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (fos != null) {
                try {
                    fos.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
            if (xssFWorkbook != null) {

                try {
                    xssFWorkbook.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }
}
