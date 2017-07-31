package com.tmc.excel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
 
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
 
public class ChangeCell {
 
    @SuppressWarnings("deprecation")
    public static void main(String[] args) {
        String fileToBeRead = "E:/test.xlsx"; // excel位置
        int coloum = 1; // 比如你要获取第1列
        try {
            XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(
                    fileToBeRead));
            XSSFSheet sheet = workbook.getSheet("Sheet2");
 
            int lines = sheet.getLastRowNum();
            int firstLine = sheet.getFirstRowNum();
            for (int i = firstLine; i <= lines; i++) {
                XSSFRow row = sheet.getRow((short) i);
                if (null == row) {
                    continue;
                } else {
                    Cell cell1 = row.getCell((short) coloum);
                    System.out.println("row.getFirstCellNum() = " + row.getFirstCellNum());
                    Cell cell = row.getCell(row.getFirstCellNum());
                    
                    if (null == cell) {
                        continue;
                    } else {
                        System.out.println(cell.getNumericCellValue());
                        int temp = (int) cell.getNumericCellValue();
                        cell.setCellValue(temp + 1);
                    }
                }
            }
            FileOutputStream out = null;
            try {
                out = new FileOutputStream(fileToBeRead);
                workbook.write(out);
            } catch (IOException e) {
                e.printStackTrace();
            } finally {
                try {
                    out.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
 
    }
 
}
