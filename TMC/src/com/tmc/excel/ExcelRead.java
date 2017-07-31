package com.tmc.excel;

import java.io.File;
import java.io.IOException;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;



public class ExcelRead {

    /**
     * @param args
     */
    public static void main(String[] args) {
        File f = new File("E:/test.xls");// E:\@英语学习\TM\Agends
        try {
            Workbook book = Workbook.getWorkbook(f);//  
            Sheet sheet = book.getSheet(0); // 获得第一个工作表对象
            for (int i = 0; i < sheet.getRows(); i++) {
                for (int j = 0; j < sheet.getColumns(); j++) {
                    Cell cell = sheet.getCell(j, i); // 获得单元格
                    System.out.print(cell.getContents() + " ");
                    //得到单元格的类型
                    //System.out.println(cell.getType());
                }
                System.out.print("\n");
            }
        } catch (BiffException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        } catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
    }

}