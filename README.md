# excel4j
[![Build Status](http://img.shields.io/travis/WillTong/excel4j.svg?branch=master)](http://img.shields.io/travis/WillTong/mybatis-helper.svg?branch=master)
[![codecov](https://codecov.io/github/WillTong/excel4j/coverage.svg?branch=master)](https://codecov.io/github/WillTong/mybatis-helper?branch=master)
[![Dependency Status](https://img.shields.io/versioneye/d/WillTong/excel4j.svg)](https://img.shields.io/versioneye/d/WillTong/mybatis-helper.svg)
[![License](https://img.shields.io/github/license/WillTong/excel4j.svg)](LICENSE)

excel导入导出的工具类，可以轻松将excel和list互转。底层使用poi。

## 使用说明
- 定义一个实体类，用来存储数据。
```java
@Excel
@Sheet("例子")
public class Example {
    @Cell("测试字符串")
    private String paramStr;
    @Cell("测试数字")
    private Integer paramInt;
    @Cell(value="测试日期",dateFormat="yyyy-MM-dd")
    private Date date;
}
```
- 导出
```java
public class Export {
    public void export(){
        List<Example> list=new ArrayList();
        FileOutputStream fout = new FileOutputStream("E:/测试文件.xls");
        fout.write(Excel4j.opsWrite().fromList(list));
        fout.close();
    }
}
```
- 导入
```java
public class Export {
    public void export(){
        byte[] data=new byte[1];
        List<Example> list = Excel4j.opsRead(data)
        .toList(Example.class);
    }
}
```