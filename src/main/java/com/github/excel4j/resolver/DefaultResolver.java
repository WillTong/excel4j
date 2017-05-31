package com.github.excel4j.resolver;

import com.github.excel4j.annotation.Cell;
import org.apache.poi.hssf.usermodel.HSSFCell;

/**
 * Created by will on 17-5-31.
 */
public interface DefaultResolver {
    void convertCell(Cell cell,Object value,HSSFCell hssfCell);
    Object convertValue(Cell cell,HSSFCell hssfCell);
    boolean support(Class<?> clazz);
}
