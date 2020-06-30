package easyExcel;

import com.alibaba.excel.EasyExcel;
import org.junit.Test;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * @program: poiAndEasyExcel
 * @description: EasyExcelWriter写操作
 * @author: cuixy
 * @create: 2020-06-30 17:40
 **/
public class EasyExcelWriter {
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



}