package com.github.excel4j.resolver;

import com.github.excel4j.annotation.Cell;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.ss.usermodel.CellType;

import java.util.Date;

/**
 * Created by will on 17-5-31.
 */
public class DateResolver implements DefaultResolver {

    @Override
    public void convertCell(Cell cell, Object value, HSSFCell hssfCell) {
//        HSSFCellStyle cellStyle = workbook.createCellStyle();
//        HSSFDataFormat format = workbook.createDataFormat();
//        cellStyle.setDataFormat(format.getFormat(cell.dateFormat()));
//        hssfCell.setCellStyle(cellStyle);
//        hssfCell.setCellValue((Date) value);
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
        return clazz==Date.class;
    }
}
