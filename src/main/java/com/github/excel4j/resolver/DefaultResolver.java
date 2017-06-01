package com.github.excel4j.resolver;

import com.github.excel4j.annotation.Cell;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCell;

/**
 * 数据类型解释器.
 * @author will
 */
public interface DefaultResolver {
    void convertCell(Cell cell, Object value, XSSFCell xssfCell,ExcelVars excelVars);
    Object convertValue(Cell cell,XSSFCell xssfCell,ExcelVars excelVars);
    boolean support(Class<?> clazz);
}
