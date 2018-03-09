package com.icexls;

/**
 * 导出Excel的工具类
 * 
 * @author iceWater
 * @version 1.0
 */
public class IceExcel {
    private String excelFileName;
    private String sheet;
    private ParserType parserType = ParserType.AUTO;
    private ExcelParser excelParser;
    private NumberType numberType = NumberType.STRING;

    /**
     * 创建一个Excel操作对象
     * 
     * @param excelFileName 操作的Excel对应的文件路径
     * @since 1.0
     */
    public IceExcel(String excelFileName) {
        this.excelFileName = excelFileName;
    }

    /**
     * 创建一个Excel操作对象
     * 
     * @param excelFileName 操作的Excel对应的文件路径
     * @param sheetName Excel对应的Sheet名称
     * @since 1.0
     */
    public IceExcel(String excelFileName, String sheetName) {
        this.excelFileName = excelFileName;
        this.sheet = sheetName;
    }

    void setParserType(ParserType parserType) {
        if (!this.parserType.equals(parserType)) {
            this.parserType = parserType;
            excelParser = null;
            init(this.sheet);
        }
    }

    /**
     * 读取Excel为String数组
     * 
     * @return 从Excel中读入的数据
     * @since 1.0
     */
    public String[][] getData() {
        init(null);
        return excelParser.getData();
    }

    /**
     * 将String二维数组导出Excel
     * 
     * @param data 需要导出Excel的数据
     * @since 1.0
     */
    public void setData(String[][] data) {
        String sheetName = "第一页";
        if (excelFileName != null) {
            int lastIndexOf = excelFileName.lastIndexOf("/");
            if (lastIndexOf != -1) {
                String fileName = excelFileName.substring(lastIndexOf + 1);
                if (fileName.endsWith(".xlsx") || fileName.endsWith(".XLSX")) {
                    fileName = fileName.substring(0, fileName.length() - 5);
                } else if (fileName.endsWith(".xls") || fileName.endsWith(".XLS")) {
                    fileName = fileName.substring(0, fileName.length() - 4);
                }
                if (fileName != null && !"".equals(fileName.trim())) {
                    sheetName = fileName;
                }
            }
        }
        init(sheetName);
        excelParser.setData(data);
    }

    private void init(String sheetName) {
        if (excelParser == null) {
            if (ParserType.AUTO.equals(parserType)) {
                if (ExcelClassUtil.hasClass("org.apache.poi.hssf.usermodel.HSSFCell")) {
                    excelParser = new PoiExcelParser();
                } else if (ExcelClassUtil.hasClass("jxl.Cell")) {
                    excelParser = new JxlExcelParser();
                } else {
                    throw new RuntimeException(
                            "没有引入poi-x.x.x.jar,你可以从下面的地址下载:http://central.maven.org/maven2/org/apache/poi/poi/3.17/poi-3.17.jar");
                }
            } else if (ParserType.JXL.equals(parserType)) {
                try {
                    excelParser = new JxlExcelParser();
                } catch (NoClassDefFoundError e) {
                    if ("jxl/read/biff/BiffException".equals(e.getMessage().trim())) {
                        throw new RuntimeException(
                                "没有引入jxl-x.x.x.jar,你可以从下面的地址下载:http://central.maven.org/maven2/net/sourceforge/jexcelapi/jxl/2.6.12/jxl-2.6.12.jar");
                    } else {
                        e.printStackTrace();
                    }
                }
            } else if (ParserType.POI.equals(parserType)) {
                try {
                    excelParser = new PoiExcelParser();
                } catch (NoClassDefFoundError e) {
                    if ("org/apache/poi/ss/usermodel/Cell".equals(e.getMessage().trim())) {
                        throw new RuntimeException(
                                "没有引入poi-x.x.x.jar,你可以从下面的地址下载:http://central.maven.org/maven2/org/apache/poi/poi/3.17/poi-3.17.jar");
                    } else {
                        e.printStackTrace();
                    }
                }
            } else {
                throw new RuntimeException("不存在的Excel实现:" + parserType);
            }
        }
        AbstractExcelParser abstractExcelParser = (AbstractExcelParser) excelParser;
        if (sheet == null) {
            sheet = sheetName;
        }
        abstractExcelParser.setExcelFileName(excelFileName);
        abstractExcelParser.setSheet(sheet);
        abstractExcelParser.setNumberType(numberType);
    }

    void setSheetName(String sheet) {
        init(sheet);
        this.sheet = sheet;
    }

    void setNumberType(NumberType numberType) {
        this.numberType = numberType;
    }

}
