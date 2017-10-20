# excel4j
[![Build Status](http://img.shields.io/travis/WillTong/excel4j.svg?branch=master)](http://img.shields.io/travis/WillTong/mybatis-helper.svg?branch=master)
[![codecov](https://codecov.io/github/WillTong/excel4j/coverage.svg?branch=master)](https://codecov.io/github/WillTong/mybatis-helper?branch=master)
[![Dependency Status](https://img.shields.io/versioneye/d/WillTong/excel4j.svg)](https://img.shields.io/versioneye/d/WillTong/mybatis-helper.svg)
[![License](https://img.shields.io/github/license/WillTong/excel4j.svg)](LICENSE)

A Java library for reading and writing Microsoft Office Excel file.Thanks to POI.

## Getting started
- Define an entity 
```java
@Excel
@Sheet("example")
public class Example {
    @Cell("testing string")
    private String paramStr;
    @Cell("testing number")
    private Integer paramInt;
    @Cell(value="testing date",dateFormat="yyyy-MM-dd")
    private Date date;
}
```
- list to xls
```java
public class Export {
    public void export(){
        List<Example> list=new ArrayList();
        FileOutputStream fout = new FileOutputStream("/test.xls");
        fout.write(Excel4j.opsWrite().fromList(list));
        fout.close();
    }
}
```
- xls to list
```java
public class Export {
    public void export(){
        byte[] data=new byte[1];
        List<Example> list = Excel4j.opsRead(data)
        .toList(Example.class);
    }
}
```