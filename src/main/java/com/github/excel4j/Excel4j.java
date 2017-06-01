package com.github.excel4j;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.ByteArrayInputStream;
import java.io.IOException;

/**
 * excel4j.
 * @author will
 */
public class Excel4j {
    private Excel4j(){}

    public static WriteOperations opsWrite(){
        return new WriteOperations(new XSSFWorkbook());
    }

    public static ReadOperations opsRead(byte[] data) throws IOException{
        return new ReadOperations(new XSSFWorkbook(new ByteArrayInputStream(data)));
    }
}
