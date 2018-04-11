package com.github.excel4j.resolver;

import com.github.excel4j.annotation.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;

import java.math.BigDecimal;

/**
 * @author will
 */
public class BigDecimalResolver implements DefaultResolver {
    @Override
    public void convertCell(Cell cell, Object value, XSSFCell xssfCell, ExcelVars excelVars) {
        xssfCell.setCellValue(((BigDecimal) value).doubleValue());
    }

    @Override
    public BigDecimal convertValue(Cell cell, XSSFCell xssfCell, ExcelVars excelVars) {
        if (xssfCell.getCellTypeEnum() == CellType.STRING){
            return new BigDecimal(xssfCell.getStringCellValue());
        }else if(xssfCell.getCellTypeEnum() == CellType.NUMERIC){
            return new BigDecimal(xssfCell.getNumericCellValue());
        }else if(xssfCell.getCellTypeEnum() == CellType.FORMULA){
            return new BigDecimal(xssfCell.getNumericCellValue());
        } else{
            return null;
        }
    }

    @Override
    public boolean support(Class<?> clazz) {
        return clazz==BigDecimal.class;
    }
}
