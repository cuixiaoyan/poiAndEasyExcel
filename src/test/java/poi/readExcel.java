package poi;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.FileInputStream;

/**
 * @program: poiAndEasyExcel
 * @description: poi读excel
 * @author: cuixy
 * @create: 2020-06-23 17:12
 **/
public class readExcel {
    //当前项目的路径。
    static String PATH = "/Users/cuixiaoyan/biancheng/utils/Java/poiAndEasyExcel/poi/";


    /**
     * 03版本
     * @throws Exception
     */
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

    /**
     * 07版本
     * @throws Exception
     */
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


}