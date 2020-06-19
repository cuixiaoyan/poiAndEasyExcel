

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

## 引入依赖

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

## 03版本

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

## 07版本



```java
@Test
public void testWrite07() throws Exception {
    // 1、创建一个工作簿 07
    Workbook workbook = new XSSFWorkbook();
    // 2、创建一个工作表
    Sheet sheet = workbook.createSheet("狂神观众统计表");
    // 3、创建一个行  （1,1）
    Row row1 = sheet.createRow(0);
    // 4、创建一个单元格
    Cell cell11 = row1.createCell(0);
    cell11.setCellValue("今日新增观众");
    // (1,2)
    Cell cell12 = row1.createCell(1);
    cell12.setCellValue(666);

    // 第二行 (2,1)
    Row row2 = sheet.createRow(1);
    Cell cell21 = row2.createCell(0);
    cell21.setCellValue("统计时间");
    // (2,2)
    Cell cell22 = row2.createCell(1);
    String time = new DateTime().toString("yyyy-MM-dd HH:mm:ss");
    cell22.setCellValue(time);

    // 生成一张表（IO 流）  03 版本就是使用 xlsx结尾！
    FileOutputStream fileOutputStream = new FileOutputStream(PATH + "狂神观众统计表07.xlsx");
    // 输出
    workbook.write(fileOutputStream);
    // 关闭流
    fileOutputStream.close();

    System.out.println("狂神观众统计表03 生成完毕！");

}
```

