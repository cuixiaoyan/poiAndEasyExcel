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