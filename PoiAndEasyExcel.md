

# EasyExcel和poi使用详解

## 引言

经常在工作或者设计毕设时，只要是有关于Excel表格的都可以用得到**poi**和**EasyExcel**，有了这两大神器之一，再也不用担心表格对你的压力了

![img](https://gitee.com/cuixiaoyan/uPic/raw/master/uPic/1905053-20200514150617200-850831803.png)

EasyExcel 是阿里巴巴开源的一个excel处理框架，**以使用简单、节省内存著称**。

EasyExcel 能大大减少占用内存的主要原因是在解析 Excel 时没有将文件数据一次性全部加载到内存中，而是从磁盘上一行行读取数据，逐个解析。

下图是 EasyExcel 和 POI 在解析Excel时的对比图。

![img](https://gitee.com/cuixiaoyan/uPic/raw/master/uPic/1905053-20200514150645200-356748885.png)

## Poi

POI是Apache软件基金会的，POI为“Poor Obfuscation Implementation”的首字母缩写，意为“简洁版的模糊实现”。
所以**POI的主要功能是可以用Java操作Microsoft Office的相关文件**，这里我们主要讲Excel

> #### 03 | 07 版本的写，就是对象不同，方法一样的！

需要注意：2003 版本和 2007 版本存在兼容性的问题！03最多只有 65535 行！

![img](https://gitee.com/cuixiaoyan/uPic/raw/master/uPic/1905053-20200514150700212-1739358610.png)

1、工作簿：

2、工作表：

3、行：

4、列：

### 引入依赖

使用junit需要放置到test文件夹下，如果要在主文件中的话，使用main方法。

![image-20200619112815537](https://gitee.com/cuixiaoyan/uPic/raw/master/uPic/image-20200619112815537.png)

```yml
 		testCompile group: 'junit', name: 'junit', version: '4.12'
    // 03(xls)
    // https://mvnrepository.com/artifact/org.apache.poi/poi
    compile group: 'org.apache.poi', name: 'poi', version: '3.17'
    // 07(xlsx)
    // https://mvnrepository.com/artifact/org.apache.poi/poi-ooxml
    compile group: 'org.apache.poi', name: 'poi-ooxml', version: '3.17'
```

### 03版本

```java
package poi;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;

import java.io.FileOutputStream;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Date;

/**
 * @program: poiAndEasyExcel
 * @description: poi常用操作
 * @author: cuixy
 * @create: 2020-06-19 10:47
 **/
public class ExcelWriter {
    //当前项目的路径。
    static String PATH = "/Users/cuixiaoyan/biancheng/utils/Java/poiAndEasyExcel/";

    @Test
    public void testWrite03() throws Exception {
        // 1、创建一个工作簿
        Workbook workbook = new HSSFWorkbook();
        // 2、创建一个工作表
        Sheet sheet = workbook.createSheet("xxx统计表");
        // 3、创建一个行  （1,1）
        Row row1 = sheet.createRow(0);
        // 4、创建一个单元格
        Cell cell11 = row1.createCell(0);
        cell11.setCellValue("今日新增观众");
        // (1,2)
        Cell cell12 = row1.createCell(1);
        cell12.setCellValue("崔笑颜");

        // 第二行 (2,1)
        Row row2 = sheet.createRow(1);
        Cell cell21 = row2.createCell(0);
        cell21.setCellValue("统计时间");
        // (2,2)
        Cell cell22 = row2.createCell(1);
        String time=LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss"));
        cell22.setCellValue(time);

        // 生成一张表（IO 流）  03 版本就是使用 xls结尾！
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "xxx统计表03.xls");
        // 输出
        workbook.write(fileOutputStream);
        // 关闭流
        fileOutputStream.close();

        System.out.println("xxx统计表03 生成完毕！");
    }

}
```

### 07版本

包路径如上。

```java
 @Test
    public void testWrite07() throws Exception {
        // 1、创建一个工作簿 07
        Workbook workbook = new XSSFWorkbook();
        // 2、创建一个工作表
        Sheet sheet = workbook.createSheet("xxx统计表");
        // 3、创建一个行  （1,1）
        Row row1 = sheet.createRow(0);
        // 4、创建一个单元格
        Cell cell11 = row1.createCell(0);
        cell11.setCellValue("今日新增观众");
        // (1,2)
        Cell cell12 = row1.createCell(1);
        cell12.setCellValue("崔笑颜");

        // 第二行 (2,1)
        Row row2 = sheet.createRow(1);
        Cell cell21 = row2.createCell(0);
        cell21.setCellValue("统计时间");
        // (2,2)
        Cell cell22 = row2.createCell(1);
        String time = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd"));
        cell22.setCellValue(time);

        // 生成一张表（IO 流）  03 版本就是使用 xlsx结尾！
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "xxx统计表07.xlsx");
        // 输出
        workbook.write(fileOutputStream);
        // 关闭流
        fileOutputStream.close();

        System.out.println("xxx统计表07 生成完毕！");

    }
```

注意对象的一个区别，文件后缀！

数据批量导入！

## 大文件

### 大文件写HSSF

缺点：最多只能处理65536行，否则会抛出异常

```java
java.lang.IllegalArgumentException: Invalid row number (65536) outside allowable range (0..65535)
```

优点：过程中写入缓存，不操作磁盘，最后一次性写入磁盘，速度快

#### 耗时：2.117

```java
@Test
public void testWrite03BigData() throws IOException {
    // 时间
    long begin = System.currentTimeMillis();

    // 创建一个薄
    Workbook workbook = new HSSFWorkbook();
    // 创建表
    Sheet sheet = workbook.createSheet();
    // 写入数据
    for (int rowNum = 0; rowNum < 65537; rowNum++) {
        Row row = sheet.createRow(rowNum);
        for (int cellNum = 0; cellNum < 10 ; cellNum++) {
            Cell cell = row.createCell(cellNum);
            cell.setCellValue(cellNum);
        }
    }
    System.out.println("over");
    FileOutputStream outputStream = new FileOutputStream(PATH + "testWrite03BigData.xls");
    workbook.write(outputStream);
    outputStream.close();
    long end = System.currentTimeMillis();
    System.out.println((double) (end-begin)/1000);
}
```

### 大文件写XSSF

缺点：写数据时速度非常慢，非常耗内存，也会发生内存溢出，如100万条

优点：可以写较大的数据量，如20万条

### 耗时：14.069

```java
@Test
public void testWrite07BigData() throws IOException {
    // 时间
    long begin = System.currentTimeMillis();

    // 创建一个薄
    Workbook workbook = new XSSFWorkbook();
    // 创建表
    Sheet sheet = workbook.createSheet();
    // 写入数据
    for (int rowNum = 0; rowNum < 100000; rowNum++) {
        Row row = sheet.createRow(rowNum);
        for (int cellNum = 0; cellNum < 10 ; cellNum++) {
            Cell cell = row.createCell(cellNum);
            cell.setCellValue(cellNum);
        }
    }
    System.out.println("over");
    FileOutputStream outputStream = new FileOutputStream(PATH + "testWrite07BigData.xlsx");
    workbook.write(outputStream);
    outputStream.close();
    long end = System.currentTimeMillis();
    System.out.println((double) (end-begin)/1000);
}
```

### 大文件写SXSSF

优点：可以写非常大的数据量，如100万条甚至更多条，写数据速度快，占用更少的内存

**注意：**

过程中会产生临时文件，需要清理临时文件

默认由100条记录被保存在内存中，如果超过这数量，则最前面的数据被写入临时文件

如果想自定义内存中数据的数量，可以使用new SXSSFWorkbook ( 数量 )

### 耗时：3.66

```java
@Test
public void testWrite07BigDataS() throws IOException {
    // 时间
    long begin = System.currentTimeMillis();

    // 创建一个薄
    Workbook workbook = new SXSSFWorkbook();
    // 创建表
    Sheet sheet = workbook.createSheet();
    // 写入数据
    for (int rowNum = 0; rowNum < 100000; rowNum++) {
        Row row = sheet.createRow(rowNum);
        for (int cellNum = 0; cellNum < 10 ; cellNum++) {
            Cell cell = row.createCell(cellNum);
            cell.setCellValue(cellNum);
        }
    }
    System.out.println("over");
    FileOutputStream outputStream = new FileOutputStream(PATH + "testWrite07BigDataS.xlsx");
    workbook.write(outputStream);
    outputStream.close();
    // 清除临时文件！
    ((SXSSFWorkbook) workbook).dispose();
    long end = System.currentTimeMillis();
    System.out.println((double) (end-begin)/1000);
}
```

SXSSFWorkbook-来至官方的解释：实现“BigGridDemo”策略的流式XSSFWorkbook版本。这允许写入非常大的文件而不会耗尽内存，因为任何时候只有可配置的行部分被保存在内存中。

请注意，仍然可能会消耗大量内存，这些内存基于您正在使用的功能，例如合并区域，注释......仍然只存储在内存中，因此如果广泛使用，可能需要大量内存。

再使用 POI的时候！内存问题 Jprofile！

## POI-Excel读

### 03版本

```java
@Test
public void testRead03() throws Exception {

    // 获取文件流
    FileInputStream inputStream = new FileInputStream(PATH + "xxx统计表03.xls");

    // 1、创建一个工作簿。 使用excel能操作的这边他都可以操作！
    Workbook workbook = new HSSFWorkbook(inputStream);
    // 2、得到表
    Sheet sheet = workbook.getSheetAt(0);
    // 3、得到行
    Row row = sheet.getRow(0);
    // 4、得到列
    Cell cell = row.getCell(1);

    // 读取值的时候，一定需要注意类型！
    // getStringCellValue 字符串类型
    System.out.println(cell.getStringCellValue());
    //System.out.println(cell.getNumericCellValue());
    inputStream.close();
}
```

### 07版本 

```java
@Test
public void testRead07() throws Exception {

    // 获取文件流
    FileInputStream inputStream = new FileInputStream(PATH + "xxx统计表07.xlsx");

    // 1、创建一个工作簿。 使用excel能操作的这边他都可以操作！
    Workbook workbook = new XSSFWorkbook(inputStream);
    // 2、得到表
    Sheet sheet = workbook.getSheetAt(0);
    // 3、得到行
    Row row = sheet.getRow(0);
    // 4、得到列
    Cell cell = row.getCell(1);

    // 读取值的时候，一定需要注意类型！
    // getStringCellValue 字符串类型
    System.out.println(cell.getStringCellValue());
    //System.out.println(cell.getNumericCellValue());
    inputStream.close();
}
```

### 读取不同的数据类型

```java
@Test
public void testCellType() throws Exception {
    // 获取文件流
    FileInputStream inputStream = new FileInputStream(PATH + "明细表.xls");

    // 创建一个工作簿。 使用excel能操作的这边他都可以操作！
    Workbook workbook = new HSSFWorkbook(inputStream);
    Sheet sheet = workbook.getSheetAt(0);
    // 获取标题内容
    Row rowTitle = sheet.getRow(0);
    if (rowTitle!=null) {
        // 一定要掌握
        int cellCount = rowTitle.getPhysicalNumberOfCells();
        for (int cellNum = 0; cellNum < cellCount; cellNum++) {
            Cell cell = rowTitle.getCell(cellNum);
            if (cell!=null){
                int cellType = cell.getCellType();
                String cellValue = cell.getStringCellValue();
                System.out.print(cellValue + " | ");
            }
        }
        System.out.println();
    }

    // 获取表中的内容
    int rowCount = sheet.getPhysicalNumberOfRows();
    for (int rowNum = 1; rowNum < rowCount ; rowNum++) {
        Row rowData = sheet.getRow(rowNum);
        if (rowData!=null){
            // 读取列
            int cellCount = rowTitle.getPhysicalNumberOfCells();
            for (int cellNum = 0; cellNum < cellCount ; cellNum++) {
                System.out.print("[" +(rowNum+1) + "-" + (cellNum+1) + "]");

                Cell cell = rowData.getCell(cellNum);
                // 匹配列的数据类型
                if (cell!=null) {
                    int cellType = cell.getCellType();
                    String cellValue = "";

                    switch (cellType) {
                        case HSSFCell.CELL_TYPE_STRING: // 字符串
                            System.out.print("【String】");
                            cellValue = cell.getStringCellValue();
                            break;
                        case HSSFCell.CELL_TYPE_BOOLEAN: // 布尔
                            System.out.print("【BOOLEAN】");
                            cellValue = String.valueOf(cell.getBooleanCellValue());
                            break;
                        case HSSFCell.CELL_TYPE_BLANK: // 空
                            System.out.print("【BLANK】");
                            break;
                        case HSSFCell.CELL_TYPE_NUMERIC: // 数字（日期、普通数字）
                            System.out.print("【NUMERIC】");
                            if (HSSFDateUtil.isCellDateFormatted(cell)){ // 日期
                                System.out.print("【日期】");
                                Date date = cell.getDateCellValue();
                                cellValue = new DateTime(date).toString("yyyy-MM-dd");
                            }else {
                                // 不是日期格式，防止数字过长！
                                System.out.print("【转换为字符串输出】");
                                cell.setCellType(HSSFCell.CELL_TYPE_STRING);
                                cellValue = cell.toString();
                            }
                            break;
                        case HSSFCell.CELL_TYPE_ERROR:
                            System.out.print("【数据类型错误】");
                            break;
                    }
                    System.out.println(cellValue);
                }
            }
        }
    }
    inputStream.close();
}
```

注意，类型转换问题；

> #### 计算公式 （了解即可！）

```java
@Test
public void testFormula() throws Exception {
    FileInputStream inputStream = new FileInputStream(PATH + "公式.xls");
    Workbook workbook = new HSSFWorkbook(inputStream);
    Sheet sheet = workbook.getSheetAt(0);

    Row row = sheet.getRow(4);
    Cell cell = row.getCell(0);

    // 拿到计算公司 eval
    FormulaEvaluator FormulaEvaluator = new HSSFFormulaEvaluator((HSSFWorkbook)workbook);

    // 输出单元格的内容
    int cellType = cell.getCellType();
    switch (cellType){
        case Cell.CELL_TYPE_FORMULA: // 公式
            String formula = cell.getCellFormula();
            System.out.println(formula);

            // 计算
            CellValue evaluate = FormulaEvaluator.evaluate(cell);
            String cellValue = evaluate.formatAsString();
            System.out.println(cellValue);
            break;
    }

}
```

# EasyExcel操作

[官方文档](https://www.yuque.com/easyexcel)

> 导入依赖

```xml
		// easyexcel
    // https://mvnrepository.com/artifact/com.alibaba/easyexcel
    compile group: 'com.alibaba', name: 'easyexcel', version: '2.2.6'

		// https://mvnrepository.com/artifact/com.alibaba/fastjson
		compile group: 'com.alibaba', name: 'fastjson', version: '1.2.72'

```

## 写入测试

[写地址](https://www.yuque.com/easyexcel/doc/read)

### 创建对象

```java
package easyExcel;

import com.alibaba.excel.annotation.ExcelIgnore;
import com.alibaba.excel.annotation.ExcelProperty;

import java.util.Date;

/**
 * @program: poiAndEasyExcel
 * @description:
 * @author: cuixy
 * @create: 2020-06-30 17:35
 **/
public class DemoData {

    @ExcelProperty("字符串标题")
    private String string;
    @ExcelProperty("日期标题")
    private Date date;
    @ExcelProperty("数字标题")
    private Double doubleData;
    /**
     * 忽略这个字段
     */
    @ExcelIgnore
    private String ignore;

    public String getString() {
        return string;
    }

    public void setString(String string) {
        this.string = string;
    }

    public Date getDate() {
        return date;
    }

    public void setDate(Date date) {
        this.date = date;
    }

    public Double getDoubleData() {
        return doubleData;
    }

    public void setDoubleData(Double doubleData) {
        this.doubleData = doubleData;
    }

    public String getIgnore() {
        return ignore;
    }

    public void setIgnore(String ignore) {
        this.ignore = ignore;
    }
}
```

### 拿到实体类里的值

```java
 //当前项目的路径。
    static String PATH = "/Users/cuixiaoyan/biancheng/utils/Java/poiAndEasyExcel/easyExcel/";

    private List<DemoData> data() {
        java.util.List<DemoData> list = new ArrayList<DemoData>();
        for (int i = 0; i < 10; i++) {
            DemoData data = new DemoData();
            data.setString("字符串" + i);
            data.setDate(new Date());
            data.setDoubleData(0.56);
            list.add(data);
        }
        return list;
    }
```

### **将list写入Excel**

```java
// 根据list 写入excel
    @Test
    public void simpleWrite() {
        // 写法1
        String fileName = PATH + "EasyTest.xlsx";
        // 这里 需要指定写用哪个class去写，然后写到第一个sheet，名字为模板 然后文件流会自动关闭
        // write (fileName, 格式类)
        // sheet (表明)
        // doWrite (数据)
        EasyExcel.write(fileName, DemoData.class).sheet("模板").doWrite(data());
    }
```

## 读取测试

[读地址](https://www.yuque.com/easyexcel/doc/read)

### 监听器

```java
package easyExcel;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.fastjson.JSON;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.ArrayList;
import java.util.List;

/**
 * @program: poiAndEasyExcel
 * @description: 读取监听器类
 * @author: cuixy
 * @create: 2020-07-05 08:49
 **/
public class DemoDataListener extends AnalysisEventListener<DemoData> {

    private static final Logger LOGGER = LoggerFactory.getLogger(DemoDataListener.class);


    private static final int BATCH_COUNT = 5;
    List<DemoData> list = new ArrayList<DemoData>();

    private DemoDAO demoDAO;

    public DemoDataListener() {
        // 这里是demo，所以随便new一个。实际使用如果到了spring,请使用下面的有参构造函数
        demoDAO = new DemoDAO();
    }

    public DemoDataListener(DemoDAO demoDAO) {
        this.demoDAO = demoDAO;
    }

    // 读取数据会执行 invoke 方法
    // DemoData 类型
    // AnalysisContext 分析上问
    @Override
    public void invoke(DemoData data, AnalysisContext context) {
        System.out.println(JSON.toJSONString(data));
        list.add(data);
        // 达到BATCH_COUNT了，需要去存储一次数据库，防止数据几万条数据在内存，容易OOM
        if (list.size() >= BATCH_COUNT) {
            saveData(); // 持久化逻辑!
            // 存储完成清理 list
            list.clear();
        }
    }

    /**
     * 所有数据解析完成了 都会来调用
     *
     * @param context
     */
    @Override
    public void doAfterAllAnalysed(AnalysisContext context) {
        // 这里也要保存数据，确保最后遗留的数据也存储到数据库
        saveData();
        LOGGER.info("所有数据解析完成！");
    }

    /**
     * 加上存储数据库
     */
    private void saveData() {
        LOGGER.info("{}条数据，开始存储数据库！", list.size());
        demoDAO.save(list);
        LOGGER.info("存储数据库成功！");
    }
}
```

### 持久层

```java
package easyExcel;

/**
 * @program: poiAndEasyExcel
 * @description: DAO存储
 * @author: cuixy
 * @create: 2020-07-05 08:50
 **/

import java.util.List;

/**
 * 假设这个是你的DAO存储。当然还要这个类让spring管理，当然你不用需要存储，也不需要这个类。
 **/
public class DemoDAO {
    public void save(List<DemoData> list) {
        // 持久化操作！
        // 如果是mybatis,尽量别直接调用多次insert,自己写一个mapper里面新增一个方法batchInsert,所有数据一次性插入
    }
}
```

### 测试代码

```java
package easyExcel;

import com.alibaba.excel.EasyExcel;
import org.junit.Test;

/**
 * @program: poiAndEasyExcel
 * @description: 测试读取方法
 * @author: cuixy
 * @create: 2020-07-05 08:53
 **/
public class simpleRead {

    static String PATH = "/Users/cuixiaoyan/biancheng/utils/Java/poiAndEasyExcel/easyExcel/";



    @Test
    public void simpleRead() {
        // 有个很重要的点 DemoDataListener 不能被spring管理，要每次读取excel都要new,然后里面用到spring可以构造方法传进去
        // 写法1：
        String fileName = PATH + "EasyTest.xlsx";
        // 这里 需要指定读用哪个class去读，然后读取第一个sheet 文件流会自动关闭

        // 重点注意读取的逻辑 DemoDataListener
        EasyExcel.read(fileName, DemoData.class, new DemoDataListener()).sheet().doRead();
    }


}
```

