import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;
import org.junit.Test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;

public class ExcelReadTest {
    static String PATH="F:\\Poi\\";//存文件的路径
    @Test
    public void testRead03() throws IOException {
        //获取文件流
        FileInputStream fileInputStream = new FileInputStream(PATH + "guyu-poi谷雨的excel写入测试.xls");
        //1.创建一个工作簿，用来接受文件流，excel中的操作再这边都可以操作
        Workbook workbook=new HSSFWorkbook(fileInputStream);
        //2.得到表
        Sheet sheet=workbook.getSheetAt(0);//也可以通过名字来获取
        //3.得到行
        Row row=sheet.getRow(0);
        //4.得到列
        Cell cell=row.getCell(0);//(1,1)
        //读取值的时候一定的要注意类型
        //getStringCellValue 获取字符串类型，或其他的，类型不一样会报错
        System.out.println(cell.getStringCellValue());
        fileInputStream.close();//关闭流
    }
    @Test
    public void testRead07() throws IOException {
        //获取文件流
        FileInputStream fileInputStream = new FileInputStream(PATH + "guyu-poi谷雨的excel写入测试.xlsx");
        //1.创建一个工作簿，用来接受文件流，excel中的操作再这边都可以操作
        Workbook workbook=new XSSFWorkbook(fileInputStream);
        //2.得到表
        Sheet sheet=workbook.getSheetAt(0);//也可以通过名字来获取
        //3.得到行
        Row row=sheet.getRow(0);
        //4.得到列
        Cell cell=row.getCell(0);//(1,1)
        //读取值的时候一定的要注意类型
        //getStringCellValue 获取字符串类型，或其他的，类型不一样会报错
        System.out.println(cell.getStringCellValue());
        fileInputStream.close();//关闭流
    }
    @Test
    public void getCellType() throws IOException {
        FileInputStream fileInputStream = new FileInputStream(PATH + "testCellType.xlsx");
        Workbook workbook=new XSSFWorkbook(fileInputStream);
        //创建一个工作簿
        Sheet sheet=workbook.getSheetAt(0);
        //获取标题的内容
        Row rowTitle=sheet.getRow(0);
        if (rowTitle!=null){
            //一定要掌握
            int cellCount = rowTitle.getPhysicalNumberOfCells();//获取这一行的列数
            for (int cellnum = 0; cellnum < cellCount; cellnum++) {
                Cell cell = rowTitle.getCell(cellnum);//获取单元格
                if(cell!=null){
                    int cellType = cell.getCellType();//获取单元格内容的类型，每个数字代表一个类型
                    String cellValue = cell.getStringCellValue();
                    System.out.print(cellValue+" | ");
                }
            }
            System.out.println();
        }
        //获取表中的内容
        int rowCount=sheet.getPhysicalNumberOfRows();//获取行数
        for (int rowNum=1;rowNum<rowCount;rowNum++){
            System.out.println("====进入这个循环====");
            Row rowData = sheet.getRow(rowNum);
            if (rowData!=null){
                //读取列
                int cellCount = rowData.getPhysicalNumberOfCells();//列数
                for (int cellNum = 0; cellNum < cellCount; cellNum++) {
                    System.out.println("["+(rowNum+1)+"-"+(cellNum+1)+"]");

                    Cell cell = rowData.getCell(cellNum);
                    //匹配列的数据类型
                    if(cell!=null){
                        int cellType = cell.getCellType();
                        String cellValue="";//最后统一转化成string类型

                        switch (cellType){
                            case HSSFCell.CELL_TYPE_STRING://字符串
                                System.out.print("【string】");
                                cellValue=cell.getStringCellValue();
                                break;
                            case HSSFCell.CELL_TYPE_BOOLEAN://布尔值
                                System.out.print("【boolean】");
                                cellValue=String.valueOf(cell.getBooleanCellValue());
                                break;
                            case HSSFCell.CELL_TYPE_BLANK://空
                                System.out.print("【blank】");
                                break;
                            case HSSFCell.CELL_TYPE_NUMERIC://数字（分为日期 普通数字）
                                System.out.print("【NUMERIC】");
                                if (HSSFDateUtil.isCellDateFormatted(cell)){//日期
                                    System.out.print("【日期】");
                                    Date date = cell.getDateCellValue();
                                    cellValue= new DateTime(date).toString("yyyy-MM-dd");//日期转化
                                }else {
                                    //不是日期格式防止数字过长
                                    System.out.print("【将数字转换层字符串输出】");
                                    cell.setCellType(HSSFCell.CELL_TYPE_STRING);
                                    cellValue=cell.toString();
                                }
                                break;
                            case HSSFCell.CELL_TYPE_ERROR://错误
                                System.out.print("【数据类型错误】");
                                break;
                        }
                        System.out.println(cellValue);
                    }
                }
            }
        }
        fileInputStream.close();//别忘了关闭流
    }
    @Test
    public void testFormula() throws IOException {
        FileInputStream fileInputStream = new FileInputStream(PATH + "公式.xlsx");
        Workbook workbook=new XSSFWorkbook(fileInputStream);
        Sheet sheet=workbook.getSheetAt(0);

        Row row=sheet.getRow(4);//有公式的单元格在第5行
        Cell cell=row.getCell(0);

        //拿到计算公式
        FormulaEvaluator formulaEvaluator=new XSSFFormulaEvaluator((XSSFWorkbook)workbook);

        //输出单元格的内容
        int cellType = cell.getCellType();
        switch (cellType){
            case Cell.CELL_TYPE_FORMULA://公式
                String formula = cell.getCellFormula();
                System.out.println(formula);

                //计算
                CellValue evaluate = formulaEvaluator.evaluate(cell);
                String cellValue = evaluate.formatAsString();
                System.out.println(cellValue);
                break;
        }
    }
}
