package com.github.excel4j.resolver;

import com.github.excel4j.annotation.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;

/**
 * Created by will on 17-5-31.
 */
public class StringResolver implements DefaultResolver {

    @Override
    public void convertCell(Cell cell, Object value, XSSFCell xssfCell, ExcelVars excelVars) {
        xssfCell.setCellValue(value.toString());
    }

    @Override
    public String convertValue(Cell cell, XSSFCell xssfCell,ExcelVars excelVars) {
        if (xssfCell.getCellTypeEnum() == CellType.STRING){
            return xssfCell.getStringCellValue();
        }else if(xssfCell.getCellTypeEnum() == CellType.NUMERIC){
            return String.valueOf(xssfCell.getNumericCellValue());
        }else{
            return null;
        }
    }

    @Override
    public boolean support(Class<?> clazz) {
        return clazz==String.class;
    }
}
