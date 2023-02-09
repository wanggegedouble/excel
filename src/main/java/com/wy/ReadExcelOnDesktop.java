package com.wy;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.joda.time.DateTime;

import java.io.FileOutputStream;
import java.io.IOException;

public class ReadExcelOnDesktop {
    public static void main(String[] args) throws IOException {
        String PATH = "C:\\Users\\huawei\\Documents\\Java_project\\excel-spring";
        String PATH_DESKTOP = "C:\\Users\\huawei\\Desktop";
        long begin = System.currentTimeMillis();
        //1：创建工作蒲
        Workbook workbook = new SXSSFWorkbook();
        //2:获取工作表：既可以根据工作表的顺序获取，也可以根据工作表的名称获取
        Sheet sheet = workbook.createSheet("wy01");
        //创建行
        for (int RowNum = 0; RowNum < 100000; RowNum++) {
            Row row1 = sheet.createRow(RowNum);
            for (int CellNum = 0; CellNum < 10; CellNum++) {
                Cell cell1 = row1.createCell(CellNum);
                cell1.setCellValue(CellNum);
            }
        }
        FileOutputStream fileOutputStream = new FileOutputStream(PATH + "wy01.xlsx");
        workbook.write(fileOutputStream);
        fileOutputStream.close();
        workbook.close();
        long end = System.currentTimeMillis();
        System.out.println("ok"+(double)(end-begin)/1000);
    }

}
