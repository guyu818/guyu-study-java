import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.junit.Test;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelWriteTest {

    @Test
    public void testWrite03() throws IOException {
        //1.创建一个工作簿
        Workbook workbook=new HSSFWorkbook();
        //2.创建一个工作表
        Sheet sheet=workbook.createSheet("sheet的名字");
        //3.创建一行
        Row row1=sheet.createRow(0);//注意第一行对应的下标为0
        //4.在这一行的基础上细分一个单元格
        Cell cell1=row1.createCell(0);//这个单元格对应（1，1）
        cell1.setCellValue("这里面设置单元格的内容，可以是整数等等");

        Cell cell2=row1.createCell(1);//(1,2)
        cell2.setCellValue(6666);

        //可以尝试多i创建几行
        Row row2=sheet.createRow(1);
        Cell cell3=row2.createCell(0);
        //使用一下joda-time这个是非常好用的
        String s = new DateTime().toString("yyyy-MM-dd HH:mm:ss");
        cell3.setCellValue("我是第二行的单元格"+s);

        //生成一张表（io流） 03版本的后缀是xls结尾
        //创建输出流
        FileOutputStream fileOutputStream=new FileOutputStream("F:\\Poi\\guyu-poi"+"谷雨的excel写入测试.xls");//抛出异常
        //输出
        workbook.write(fileOutputStream);//抛出异常
        //关闭流
        fileOutputStream.close();

        System.out.println("表格生成完毕！");
    }
    @Test
    public void testWrite07() throws IOException {
        //1.创建一个工作簿
        Workbook workbook=new XSSFWorkbook();//07和03对象不同的区别
        //2.创建一个工作表
        Sheet sheet=workbook.createSheet("sheet的名字");
        //3.创建一行
        Row row1=sheet.createRow(0);//注意第一行对应的下标为0
        //4.在这一行的基础上细分一个单元格
        Cell cell1=row1.createCell(0);//这个单元格对应（1，1）
        cell1.setCellValue("这里面设置单元格的内容，可以是整数等等");

        Cell cell2=row1.createCell(1);//(1,2)
        cell2.setCellValue(6666);

        //可以尝试多i创建几行
        Row row2=sheet.createRow(1);
        Cell cell3=row2.createCell(0);
        //使用一下joda-time这个是非常好用的
        String s = new DateTime().toString("yyyy-MM-dd HH:mm:ss");
        cell3.setCellValue("我是第二行的单元格"+s);

        //生成一张表（io流） 07版本的后缀是xlsx结尾
        //创建输出流
        FileOutputStream fileOutputStream=new FileOutputStream("F:\\Poi\\guyu-poi"+"谷雨的excel写入测试.xlsx");//抛出异常
        //输出
        workbook.write(fileOutputStream);//抛出异常
        //关闭流
        fileOutputStream.close();

        System.out.println("表格生成完毕！");
    }
    @Test
    public void testWrite03BigData() throws IOException {
        //起始时间
        long begin=System.currentTimeMillis();
        //创建簿
        Workbook workbook=new HSSFWorkbook();
        //创建表
        Sheet sheet=workbook.createSheet();
        //写入数据
        for (int i = 0; i < 65536; i++) {
            Row row=sheet.createRow(i);
            for (int i1 = 0; i1 < 10; i1++) {
                Cell cell=row.createCell(i1);
                cell.setCellValue(i1);
            }
        }
        System.out.println("over");
        //创建输出流
        FileOutputStream fileOutputStream = new FileOutputStream("F:\\Poi\\guyu-poi" + "03bigdata.xls");
        workbook.write(fileOutputStream);//输出
        fileOutputStream.close();//关闭
        long end=System.currentTimeMillis();
        System.out.println((double)(end-begin)/1000);
        //测试时间2.059秒
    }

    //速度慢，存储量大，如何优化速度，缓存
    @Test
    public void testWrite07BigData() throws IOException {
        //起始时间
        long begin=System.currentTimeMillis();
        //创建簿
        Workbook workbook=new XSSFWorkbook();
        //创建表
        Sheet sheet=workbook.createSheet();
        //写入数据
        for (int i = 0; i < 65536; i++) {//这里改大点就可以了最多可以存100多万行数据
            Row row=sheet.createRow(i);
            for (int i1 = 0; i1 < 10; i1++) {
                Cell cell=row.createCell(i1);
                cell.setCellValue(i1);
            }
        }
        System.out.println("over");
        //创建输出流
        FileOutputStream fileOutputStream = new FileOutputStream("F:\\Poi\\guyu-poi" + "07bigdata.xlsx");
        workbook.write(fileOutputStream);//输出
        fileOutputStream.close();//关闭
        long end=System.currentTimeMillis();
        System.out.println((double)(end-begin)/1000);
        //测试时间19.002秒
    }
    @Test
    public void testWrite07BigDataS() throws IOException {
        //起始时间
        long begin=System.currentTimeMillis();
        //创建簿
        Workbook workbook=new SXSSFWorkbook();//这里变成加速的07了
        //创建表
        Sheet sheet=workbook.createSheet();
        //写入数据
        for (int i = 0; i < 65536; i++) {//这里改大点就可以了最多可以存100多万行数据
            Row row=sheet.createRow(i);
            for (int i1 = 0; i1 < 10; i1++) {
                Cell cell=row.createCell(i1);
                cell.setCellValue(i1);
            }
        }
        System.out.println("over");
        //创建输出流
        FileOutputStream fileOutputStream = new FileOutputStream("F:\\Poi\\guyu-poi" + "07bigdataS.xlsx");
        workbook.write(fileOutputStream);//输出
        fileOutputStream.close();//关闭
        //清除临时文件,记住
        ((SXSSFWorkbook) workbook).dispose();
        long end=System.currentTimeMillis();
        System.out.println((double)(end-begin)/1000);
        //测试时间2.839秒
    }
}
