package com.github.excel4j;

import com.github.excel4j.annotation.Cell;
import com.github.excel4j.annotation.Excel;
import com.github.excel4j.resolver.*;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

/**
 * read operation.
 * @author will
 */
public class ReadOperations {
    private ExcelVars excelVars;
    private XSSFWorkbook workbook;
    private List<DefaultResolver> resolverList;

    protected ReadOperations(XSSFWorkbook workbook){
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

    public <T> List<T> toList(Class<T> clazz) throws IllegalAccessException, InstantiationException {
        List<T> list = new ArrayList<>();
        Annotation excel = clazz.getAnnotation(Excel.class);
        if (excel == null) {
            throw new NullPointerException("该类无法使用Excel4j处理！");
        }
        if (workbook.getSheetAt(0) == null) {
            throw new NullPointerException("没有可供处理的sheet页！");
        }
        XSSFSheet xssfSheet = workbook.getSheetAt(0);
        //通过表头获得序列
        int[] fieldIndexes = new int[clazz.getDeclaredFields().length];
        if (xssfSheet.getLastRowNum() == 0) {
            throw new NullPointerException("sheet页为空！");
        }
        for (int fieldIndex = 0; fieldIndex < clazz.getDeclaredFields().length; fieldIndex++) {
            boolean isHas = false;
            for (int colIndex = 0; colIndex < xssfSheet.getRow(0).getLastCellNum(); colIndex++) {
                Cell cell = clazz.getDeclaredFields()[fieldIndex].getAnnotation(Cell.class);
                if (cell != null && cell.value().equals(xssfSheet.getRow(0).getCell(colIndex).getStringCellValue())) {
                    fieldIndexes[fieldIndex] = colIndex;
                    isHas = true;
                    break;
                }
            }
            if (!isHas) {
                fieldIndexes[fieldIndex] = -1;
            }
        }
        //获取数据
        for (int rowIndex = 1; rowIndex <= workbook.getSheetAt(0).getLastRowNum(); rowIndex++) {
            XSSFRow xssfRow = workbook.getSheetAt(0).getRow(rowIndex);
            Object object = clazz.newInstance();
            for (int fieldIndex = 0; fieldIndex < fieldIndexes.length; fieldIndex++) {
                if (fieldIndexes[fieldIndex] != -1) {
                    Field field = clazz.getDeclaredFields()[fieldIndex];
                    field.setAccessible(true);
                    XSSFCell xssfCell = xssfRow.getCell(fieldIndexes[fieldIndex]);
                    if (null == xssfCell) {
                    } else {
                        Cell cell=field.getAnnotation(Cell.class);
                        if (xssfCell.getCellTypeEnum() == CellType.BLANK) {

                        } else{
                            for(DefaultResolver resolver : resolverList){
                                if(resolver.support(field.getType())){
                                    field.set(object,resolver.convertValue(cell,xssfCell,excelVars));
                                    break;
                                }
                            }
                        }
                    }
                }
            }
            list.add((T)object);
        }
        return list;
    }

}
