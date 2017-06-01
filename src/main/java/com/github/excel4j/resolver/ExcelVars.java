package com.github.excel4j.resolver;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.Map;

/**
 * Created by will on 2017/6/1.
 */
public class ExcelVars {
    private XSSFWorkbook workbook;
    private Map<String,CellStyle> styleMap;

    public XSSFWorkbook getWorkbook() {
        return workbook;
    }

    public void setWorkbook(XSSFWorkbook workbook) {
        this.workbook = workbook;
    }

    public Map<String, CellStyle> getStyleMap() {
        return styleMap;
    }

    public void setStyleMap(Map<String, CellStyle> styleMap) {
        this.styleMap = styleMap;
    }
}
