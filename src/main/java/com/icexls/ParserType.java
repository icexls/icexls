package com.icexls;
/**
 * 导出Excel使用的API操作类型
 * AUTO - 根据存在的jar包自动选择
 * JXL  - 使用jxl api
 * POI  - 使用POI api
 * 
 * @author iceWater
 * @version 1.0
 */
public enum ParserType {
    AUTO, JXL, POI
}
