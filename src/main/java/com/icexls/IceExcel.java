package com.icexls;

/**
 * 导出Excel的工具类
 * 
 * @author iceWater
 * @date 2017-04-08
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
     * @param excelFileName
     *            操作的Excel对应的文件路径
     */
    public IceExcel(String excelFileName) {
        this.excelFileName = excelFileName;
    }

    /**
     * 创建一个Excel操作对象
     * 
     * @param excelFileName
     *            操作的Excel对应的文件路径
     * @param sheetName
     *            Excel对应的Sheet名称
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
     */
    public String[][] getData() {
        init(null);
        return excelParser.getData();
    }

    /**
     * 将String二维数组导出Excel
     * 
     * @param data
     *            需要导出Excel的数据
     */
    public void setData(String[][] data) {
        init("第一页");
        excelParser.setData(data);
    }

    private void init(String sheetName) {
        if (excelParser == null) {
            if (ParserType.AUTO.equals(parserType)) {
                if (ExcelClassUtil.hasClass("jxl.Cell")) {
                    excelParser = new JxlExcelParser();
                } else {
                    excelParser = new PoiExcelParser();
                }
            } else if (ParserType.JXL.equals(parserType)) {
                excelParser = new JxlExcelParser();
            } else if (ParserType.POI.equals(parserType)) {
                excelParser = new PoiExcelParser();
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
