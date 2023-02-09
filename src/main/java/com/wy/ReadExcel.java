package com.wy;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

public class ReadExcel {
    public static void main(String[] args) throws IOException {
        String PATH = "C:\\Users\\huawei\\";
        FileInputStream fileInputStream = new FileInputStream(PATH+"Desktopwy.xlsx");
        Workbook workbook = new XSSFWorkbook(fileInputStream);
        Sheet sheet = workbook.getSheetAt(0);
        Row row = sheet.getRow(0);
        for (int CellNum = 0; CellNum < 10; CellNum++) {
            Cell cell = row.getCell(CellNum);
            System.out.println((int) cell.getNumericCellValue());
        }
        fileInputStream.close();
    }
}
