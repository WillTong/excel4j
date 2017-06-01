package com.github.excel4j;

import com.github.excel4j.annotation.Cell;
import com.github.excel4j.annotation.Excel;
import com.github.excel4j.annotation.Sheet;
import com.github.excel4j.resolver.*;
import org.apache.poi.xssf.usermodel.*;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

/**
 * write operation.
 * @author will
 */
public class WriteOperations {
    private ExcelVars excelVars;
    private XSSFWorkbook workbook;
    private List<DefaultResolver> resolverList;

    protected WriteOperations(XSSFWorkbook workbook){
        this.workbook=workbook;
        excelVars=new ExcelVars();
        excelVars.setStyleMap(new HashMap<>());
        excelVars.setWorkbook(workbook);
        resolverList=new ArrayList<>();
        resolverList.add(new IntegerResolver());
        resolverList.add(new DateResolver());
        resolverList.add(new BigDecimalResolver());
        resolverList.add(new StringResolver());
    }

    public byte[] fromList(List list) throws IllegalAccessException, IOException {
        if (list == null || list.size() == 0) {
            throw new NullPointerException("没有数据！");
        }
        Object object = list.get(0);
        Class clazz=object.getClass();
        Annotation excelAnno = clazz.getAnnotation(Excel.class);
        if (excelAnno == null) {
            throw new NullPointerException("该类无法使用Excel4j处理！");
        }
        Annotation sheetAnno =  clazz.getAnnotation(Sheet.class);
        XSSFSheet xssfSheet;
        if (sheetAnno == null) {
            xssfSheet = workbook.createSheet();
        } else {
            xssfSheet = workbook.createSheet(((Sheet)sheetAnno).value());
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
                field.setAccessible(true);
                Cell cell = field.getAnnotation(Cell.class);
                if (cell != null) {
                    Object value =field.get(data);
                    XSSFCell xssfCell = xssfRowContent.createCell(colIndex);
                    if(value==null){
                        xssfCell.setCellValue("");
                    }else{
                        for(DefaultResolver resolver : resolverList){
                            if(resolver.support(value.getClass())){
                                resolver.convertCell(cell,value,xssfCell,excelVars);
                                break;
                            }
                        }
                    }
                }
            }
        }
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        workbook.write(outputStream);
        return outputStream.toByteArray();
    }

}
