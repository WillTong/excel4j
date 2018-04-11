package com.github.excel4j.resolver;

import com.github.excel4j.annotation.Cell;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;

/**
 * @author will
 */
public class IntegerResolver implements DefaultResolver {

    @Override
    public void convertCell(Cell cell, Object value, XSSFCell xssfCell, ExcelVars excelVars) {
        xssfCell.setCellValue((Integer) value);
    }

    @Override
    public Integer convertValue(Cell cell, XSSFCell xssfCell,ExcelVars excelVars) {
        if (xssfCell.getCellTypeEnum() == CellType.STRING){
            return Integer.parseInt(xssfCell.getStringCellValue());
        }else if(xssfCell.getCellTypeEnum() == CellType.NUMERIC){
            return (int)xssfCell.getNumericCellValue();
        }else if(xssfCell.getCellTypeEnum() == CellType.FORMULA){
            return (int)xssfCell.getNumericCellValue();
        } else{
            return null;
        }
    }

    @Override
    public boolean support(Class<?> clazz) {
        return clazz==Integer.class||clazz==int.class;
    }
}
