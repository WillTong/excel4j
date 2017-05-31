package com.github.excel4j.resolver;

import com.github.excel4j.annotation.Cell;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.ss.usermodel.CellType;

/**
 * Created by will on 17-5-31.
 */
public class IntegerResolver implements DefaultResolver {

    @Override
    public void convertCell(Cell cell, Object value, HSSFCell hssfCell) {
        hssfCell.setCellValue((Integer) value);
    }

    @Override
    public Integer convertValue(Cell cell, HSSFCell hssfCell) {
        if (hssfCell.getCellTypeEnum() == CellType.STRING){
            return Integer.parseInt(hssfCell.getStringCellValue());
        }else if(hssfCell.getCellTypeEnum() == CellType.NUMERIC){
            return (int)hssfCell.getNumericCellValue();
        }else{
            return null;
        }
    }

    @Override
    public boolean support(Class<?> clazz) {
        return clazz==Integer.class||clazz==int.class;
    }
}
