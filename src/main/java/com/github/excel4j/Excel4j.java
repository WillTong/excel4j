package com.github.excel4j;

import com.github.excel4j.annotation.Cell;
import com.github.excel4j.annotation.Excel;
import com.github.excel4j.annotation.Sheet;
import com.github.excel4j.exception.ExcelException;
import com.github.excel4j.resolver.DefaultResolver;
import org.apache.poi.xssf.usermodel.*;

import java.io.ByteArrayOutputStream;
import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * Created by will on 17-5-31.
 */
public class Excel4j {
    private XSSFWorkbook workbook;
    private List<DefaultResolver> resolverList;


    private Excel4j(){
        this.workbook=new XSSFWorkbook();
    }

    public byte[] save(List list){
        try{
            if (list == null || list.size() == 0) {
                throw new ExcelException("没有数据！");
            }
            Object object = list.get(0);
            Class clazz=object.getClass();
            Annotation excelAnno = clazz.getAnnotation(Excel.class);
            if (excelAnno == null) {
                throw new ExcelException("该类无法使用Excel4j处理！");
            }
            Sheet sheetAnno = clazz.<Sheet>getAnnotation(Sheet.class);
            XSSFSheet xssfSheet;
            if (sheetAnno == null) {
                xssfSheet = workbook.createSheet();
            } else {
                xssfSheet = workbook.createSheet(sheetAnno.value());
            }
            //表头
            List<Field> fieldList=new ArrayList<>();
            XSSFRow xssfRowHeader = xssfSheet.createRow(0);
            for (int i=0;i<clazz.getDeclaredFields().length;i++) {
                Field field=clazz.getDeclaredFields()[i];
                Cell cell = field.getAnnotation(Cell.class);
                if (cell != null) {
                    XSSFCell xssfCell = xssfRowHeader.createCell(i);
                    xssfCell.setCellValue(cell.value());
                    xssfSheet.setColumnWidth(i, cell.width());
                    fieldList.add(field);
                }
            }
            //内容
            XSSFCellStyle cellStyle = workbook.createCellStyle();
            for (int rowIndex = 0; rowIndex < list.size(); rowIndex++) {
                Object data = list.get(rowIndex);
                XSSFRow xssfRowContent = xssfSheet.createRow(rowIndex + 1);
                for (int colIndex=0;colIndex<fieldList.size();colIndex++) {
                    Field field=fieldList.get(colIndex);
                    Cell cell = field.getAnnotation(Cell.class);
                    if (cell != null) {
                        Object value =field.get(data);
                        XSSFCell xssfCell = xssfRowContent.createCell(colIndex);
                        if (value == null) {
                            xssfCell.setCellValue("");
                        } else if (value instanceof Date) {
                            XSSFDataFormat format = workbook.createDataFormat();
                            cellStyle.setDataFormat(format.getFormat(cell.dateFormat()));
                            xssfCell.setCellStyle(cellStyle);
                            xssfCell.setCellValue((Date) value);
                        } else if (value instanceof Integer) {
                            xssfCell.setCellValue((Integer) value);
                        } else if (value instanceof BigDecimal) {
                            xssfCell.setCellValue(((BigDecimal) value).doubleValue());
                        } else {
                            xssfCell.setCellValue(value.toString());
                        }
                        colIndex++;
                    }
                }
            }
            ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
            workbook.write(outputStream);
            return outputStream.toByteArray();
        }catch(Exception e){

        }
        return null;
    }
}
