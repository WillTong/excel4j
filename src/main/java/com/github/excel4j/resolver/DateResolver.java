package com.github.excel4j.resolver;

import com.github.excel4j.annotation.Cell;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;

import java.util.Date;

/**
 * @author will
 */
public class DateResolver implements DefaultResolver {

    @Override
    public void convertCell(Cell cell, Object value, XSSFCell xssfCell, ExcelVars excelVars) {
        if(excelVars.getStyleMap().containsKey(cell.dateFormat())){
            xssfCell.setCellStyle(excelVars.getStyleMap().get(cell.dateFormat()));
            xssfCell.setCellValue((Date) value);
        }else{
            XSSFCellStyle cellStyle = excelVars.getWorkbook().createCellStyle();
            XSSFDataFormat format = excelVars.getWorkbook().createDataFormat();
            cellStyle.setDataFormat(format.getFormat(cell.dateFormat()));
            excelVars.getStyleMap().put(cell.dateFormat(),cellStyle);
            xssfCell.setCellStyle(cellStyle);
            xssfCell.setCellValue((Date) value);
        }
    }

    @Override
    public Date convertValue(Cell cell,XSSFCell xssfCell, ExcelVars excelVars) {
        if (xssfCell.getCellTypeEnum() == CellType.STRING){
            return null;
        }else if(xssfCell.getCellTypeEnum() == CellType.NUMERIC){
            return xssfCell.getDateCellValue();
        }else{
            return null;
        }
    }

    @Override
    public boolean support(Class<?> clazz) {
        return clazz==Date.class;
    }
}
