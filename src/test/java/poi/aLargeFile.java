package poi;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.FileOutputStream;
import java.io.IOException;

/**
 * @program: poiAndEasyExcel
 * @description: poi大文件读写
 * @author: cuixy
 * @create: 2020-06-22 14:17
 **/
public class aLargeFile {
    static String PATH = "/Users/cuixiaoyan/biancheng/utils/Java/poiAndEasyExcel/poi/";

    /**
     * 耗时：2.117 秒
     * 最多：65536 行
     *
     * @throws IOException
     */
    @Test
    public void testWrite03BigData() throws IOException {
        //耗时
        long begin = System.currentTimeMillis();
        //创建一个簿
        HSSFWorkbook workbook = new HSSFWorkbook();
        //创建表
        HSSFSheet sheet = workbook.createSheet();
        //写入数据
        for (int rowNum = 0; rowNum < 65536; rowNum++) {
            HSSFRow row = sheet.createRow(rowNum);
            for (int cellNum = 0; cellNum < 10; cellNum++) {
                HSSFCell cell = row.createCell(cellNum);
                cell.setCellValue(cellNum);
            }

        }
        System.out.println("over");
        FileOutputStream outputStream = new FileOutputStream(PATH + "testWrite03BigData.xls");
        workbook.write(outputStream);
        outputStream.close();
        long end = System.currentTimeMillis();
        System.out.println((double) (end - begin) / 1000);

    }

    /**
     * 耗时：14.069 秒
     * 行数：100000 行
     *
     * @throws IOException
     */
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
            for (int cellNum = 0; cellNum < 10; cellNum++) {
                Cell cell = row.createCell(cellNum);
                cell.setCellValue(cellNum);
            }
        }
        System.out.println("over");
        FileOutputStream outputStream = new FileOutputStream(PATH + "testWrite07BigData.xlsx");
        workbook.write(outputStream);
        outputStream.close();
        long end = System.currentTimeMillis();
        System.out.println((double) (end - begin) / 1000);
    }


    /**
     * 耗时：3.66 秒
     * 行数：100000 行
     *
     * @throws IOException
     */
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
            for (int cellNum = 0; cellNum < 10; cellNum++) {
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
        System.out.println((double) (end - begin) / 1000);
    }


}