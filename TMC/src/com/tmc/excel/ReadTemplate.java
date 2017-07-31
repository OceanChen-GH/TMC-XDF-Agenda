package com.tmc.excel;

import java.io.BufferedReader;  
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;  
import java.io.IOException;  
import java.io.InputStreamReader;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook; 

public class ReadTemplate {
	
	/* map:例子
	{Word of the day=Thanksgiving, TIMER(agenda)=Gen&Victor, Ah Counter=Ocean, Warm up Session=Ocean, SAA=Tracy, TOM(review)=Alwin, IE1=Harry, IE3=Ivy Chen（CL1）, Greeter=Ivy, IE2=Carol Fang , GRAMMARIAN=Melody, SPEAKER2=Tracy(CC1), SPEAKER1=Elena(CC5), SPEAKER3=Harry(CC7), GE=Jamie}
	
	注意：map中TIMER(agenda) change to TIMER; TOM(review) change to TOM
	{Word of the day=merry, IE1For=IE For Angelia , TTE=Elena, IE3For=IE For Tracy , IE2For=IE For Lancy , TIMER=Ivy, Ah Counter=Paul, TTM=Angelia, Title1= Colorful Life Makes an Outgoing Girl, Title3= Hey，it's me, Title2= Story of My Name, Warm up Session=Ivy, SAA=Judy, TOM=Jason, IE1=Tina, IE3=, Greeter=Elena, IE2=, Theme=Christmas Eve, Ice Breaking Eve, GRAMMARIAN=Tina, SPEAKER2=Lancy (CC1), SPEAKER1=Angelia (CC1), SPEAKER3=Tracy (CC1), GE=Alwin}
	*/
	private static Map<String, String> rolesMap = new HashMap<String, String>();
	
	public static void main(String[] args) throws IOException {
		// 调试
//		readF1("E:/@英语学习/TMC/@Agenda-Tools/AgendaGenerator-XDF/roles-xxth.txt");
//		changeCell("E:/@英语学习/TMC/@Agenda-Tools/AgendaGenerator-XDF/agenda-xxth.xlsx");
		
		// 打包
		// put roles into rolesMap
		readF1("./roles-xxth.txt");
		// modify Excel according to rolesMap
		changeCell("./agenda-xxth.xlsx");
		
//		readF1(args[0]);
//		changeCell(args[1]);
	}
	
    public static final void readF1(String filePath) throws IOException {  
        BufferedReader br = new BufferedReader(new InputStreamReader(  
                new FileInputStream(filePath), "UTF-8"));  
  
        for (String line = br.readLine(); line != null; line = br.readLine()) {  
            if (line.length() >= 17) {
            	if (line.substring(0, 16).equalsIgnoreCase("Word of the day:")) {// first line is "Word of the day:"
	            	System.out.println(line.substring(17));
	            	rolesMap.put(line.substring(0, 15),	line.substring(16));
	            	continue;
            	}
            }
            if (line.length() >= 7 ) {
            	if (line.substring(0, 5).equalsIgnoreCase("Theme")) {
            		System.out.println(line.substring(7));
            		rolesMap.put(line.substring(0, 5),	line.substring(7));
            		continue;
            	}	
            }
            if (line.length() >= 5 ) {
            	if (line.substring(0, 3).equalsIgnoreCase("SAA")) {
            		System.out.println(line.substring(5));
            		rolesMap.put(line.substring(0, 3),	line.substring(5));
            		continue;
            	}	
            }
            if (line.length() >= 9) {
            	if (line.substring(0, 7).equalsIgnoreCase("Greeter")) {
            		System.out.println(line.substring(9));
                	rolesMap.put(line.substring(0, 7),	line.substring(9));
                	continue;
            	}
            }
            if (line.length() >= 5) {
            	if (line.substring(0, 4).equalsIgnoreCase("TOM—")
            			|| line.substring(0, 4).equalsIgnoreCase("TOM-")) {
                	System.out.println(line.substring(5));
                	rolesMap.put(line.substring(0, 3),	line.substring(5));
                	continue;
            	}
            }
            if (line.length() >= 13) {
            	if (line.substring(0, 11).equalsIgnoreCase("TOM(review)")) {
                	System.out.println(line.substring(13));
                	rolesMap.put(line.substring(0, 3),	line.substring(13));
                	continue;
            	}
            }
            if (line.length() >= 4) {
            	if (line.substring(0, 2).equalsIgnoreCase("GE")) {
                	System.out.println(line.substring(4));
                	rolesMap.put(line.substring(0, 2),	line.substring(4));
                	continue;
            	}
            }
            if (line.length() >= 12) {
            	if (line.substring(0, 10).equalsIgnoreCase("Ah Counter")) {
                	System.out.println(line.substring(12));
                	rolesMap.put(line.substring(0, 10),	line.substring(12));
                	continue;
            	}
            }
            if (line.length() >= 6) {
            	if(line.substring(0, 6).equalsIgnoreCase("TIME—")
            			|| line.substring(0, 6).equalsIgnoreCase("TIME-")) {
                	System.out.println(line.substring(6));
                	rolesMap.put(line.substring(0, 5),	line.substring(7));
                	continue;
            	}
            }
            if (line.length() >= 15) {
            	if(line.substring(0, 13).equalsIgnoreCase("TIMER(agenda)")) {
                	System.out.println(line.substring(15));
                	rolesMap.put(line.substring(0, 5),	line.substring(15));
                	continue;
            	}
            }
            if (line.length() >= 12) {
            	if (line.substring(0, 10).equalsIgnoreCase("GRAMMARIAN")) {
                	System.out.println(line.substring(12));
                	rolesMap.put(line.substring(0, 10),	line.substring(12));
                	continue;
            	}
            }
            if (line.length() >= 17) {
            	if (line.substring(0, 15).equalsIgnoreCase("Warm up Session")) {
                	System.out.println(line.substring(17));
                	rolesMap.put(line.substring(0, 15),	line.substring(17));
                	continue;
            	}
            }
            if (line.length() >= 5) {
            	if (line.substring(0, 3).equalsIgnoreCase("TTM")) {
                	System.out.println(line.substring(5));
                	rolesMap.put(line.substring(0, 3),	line.substring(5));
                	continue;
            	}
            }
            if (line.length() >= 5) {
            	if (line.substring(0, 3).equalsIgnoreCase("TTE")) {
                	System.out.println(line.substring(5));
                	rolesMap.put(line.substring(0, 3),	line.substring(5));
                	continue;
            	}
            }
            if (line.length() >= 10) {
            	if (line.substring(0, 8).equalsIgnoreCase("SPEAKER1")
            			|| line.substring(0, 8).equalsIgnoreCase("SPEAKER2")
            			|| line.substring(0, 8).equalsIgnoreCase("SPEAKER3")) {
            		System.out.println(line.substring(10));
            		int nameEnd;
            		if (line.contains("(")) {
            			nameEnd = line.indexOf("(");
            		} else if (line.contains("（")) {
            			nameEnd = line.indexOf("（");
            		} else {
            			nameEnd = 11;
            		}
            		
            		int nameCCEnd;
            		if (line.contains(")")) {
            			nameCCEnd = line.indexOf(")");
            		} else if (line.contains("）")) {
            			nameCCEnd = line.indexOf("）");
            		} else {
            			nameCCEnd = 14;// 11+3
            		}
            		
                	rolesMap.put(line.substring(0, 8),	line.substring(10, nameCCEnd+1));
                	
                	switch (line.substring(0, 8)) {
                	case "SPEAKER1":
                		rolesMap.put("Title1", line.substring(nameCCEnd+1));
                		rolesMap.put("IE1For", "IE For " + line.substring(10, nameEnd));
                		break;
                	case "SPEAKER2":
                		rolesMap.put("Title2", line.substring(nameCCEnd+1));
                		rolesMap.put("IE2For", "IE For " + line.substring(10, nameEnd));
                		break;
                	case "SPEAKER3":
                		rolesMap.put("Title3", line.substring(nameCCEnd+1));
                		rolesMap.put("IE3For", "IE For " + line.substring(10, nameEnd));
                		break;
                	}
                	
                	continue;
            	}	
            }
            if (line.length() >= 5) {
            	if (line.substring(0, 3).equalsIgnoreCase("IE1")
            			|| line.substring(0, 3).equalsIgnoreCase("IE2")
            			|| line.substring(0, 3).equalsIgnoreCase("IE3")) {
            		System.out.println(line.substring(5));
                	rolesMap.put(line.substring(0, 3),	line.substring(5));
                	continue;
            	}
            } 
        }
        br.close();  
  
    }  
  
    public static final void readF2(String filePath) throws IOException {  
        FileReader fr = new FileReader(filePath);  
        BufferedReader bufferedreader = new BufferedReader(fr);  
        String instring;  
        while ((instring = bufferedreader.readLine().trim()) != null) {  
            if (0 != instring.length()) {  
                System.out.println(instring);  
            }  
        }  
        fr.close();  
    }
    
    public static final void changeCell(String excelFile){
        String fileToBeRead = excelFile;// "E:/test.xlsx"; // excel位置
        try {
            XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(
                    fileToBeRead));
            XSSFSheet sheet = workbook.getSheet("Sheet1");
//            int firstLine = 3;// role从第4行开始
//            int lines = sheet.getLastRowNum();
            // 变量rolesMap中的元素
            for (Map.Entry<String, String> entry : rolesMap.entrySet()) {
            	XSSFRow row = null;
            	Cell cell = null;
            	Cell cell2 = null;// Timer等可能需要修改两行
            	switch (entry.getKey().toUpperCase()) {
            	case "THEME":
            		row = sheet.getRow(6);
            		cell = row.getCell(1);
            		cell.setCellValue(entry.getValue());
            		
            		row = sheet.getRow(23);
            		cell = row.getCell(3);
            		cell.setCellValue(entry.getValue());
            		break;
            	case "WORD OF THE DAY":
            		row = sheet.getRow(8);
            		cell = row.getCell(1);
            		cell.setCellValue(entry.getValue());
            		break;
            	case "SAA":
            		row = sheet.getRow(3);
            		cell = row.getCell(5);// F列，第6列，从0计数
            		cell.setCellValue(entry.getValue());
            		break;
            	case "TOM":
            		row = sheet.getRow(5);
            		cell = row.getCell(5);
            		cell.setCellValue(entry.getValue());
            		
            		row = sheet.getRow(32);
            		cell2 = row.getCell(5);
            		cell2.setCellValue(entry.getValue());
            		break;
            	case "WARM UP SESSION":
            		row = sheet.getRow(6);
            		cell = row.getCell(5);
            		cell.setCellValue(entry.getValue());
            		break;
            	case "GREETER":
            		row = sheet.getRow(7);
            		cell = row.getCell(5);
            		cell.setCellValue(entry.getValue());
            		break;
            	case "TIMER":
            		row = sheet.getRow(11);
            		cell = row.getCell(5);
            		cell.setCellValue(entry.getValue());
            		
            		row = sheet.getRow(29);
            		cell2 = row.getCell(5);
            		cell2.setCellValue(entry.getValue());
            		break;
            	case "AH COUNTER":
            		row = sheet.getRow(9);
            		cell = row.getCell(5);
            		cell.setCellValue(entry.getValue());
            		
            		row = sheet.getRow(27);
            		cell2 = row.getCell(5);
            		cell2.setCellValue(entry.getValue());
            		break;
            	case "GE":
            		row = sheet.getRow(8);
            		cell = row.getCell(5);
            		cell.setCellValue(entry.getValue());
            		
            		row = sheet.getRow(30);
            		cell2 = row.getCell(5);
            		cell2.setCellValue(entry.getValue());
            		break;
            	case "GRAMMARIAN":
            		row = sheet.getRow(10);
            		cell = row.getCell(5);
            		cell.setCellValue(entry.getValue());
            		
            		row = sheet.getRow(28);
            		cell2 = row.getCell(5);
            		cell2.setCellValue(entry.getValue());
            		break;
            	case "SPEAKER1":
            		row = sheet.getRow(13);
            		cell = row.getCell(5);
            		cell.setCellValue(entry.getValue());
            		break;
            	case "TITLE1":
            		row = sheet.getRow(13);
            		cell = row.getCell(3);
            		cell.setCellValue(entry.getValue());
            		break;
            	case "SPEAKER2":
            		row = sheet.getRow(14);
            		cell = row.getCell(5);
            		cell.setCellValue(entry.getValue());
            		break;
            	case "TITLE2":
            		row = sheet.getRow(14);
            		cell = row.getCell(3);
            		cell.setCellValue(entry.getValue());
            		break;
            	case "SPEAKER3":
            		row = sheet.getRow(15);
            		cell = row.getCell(5);
            		cell.setCellValue(entry.getValue());
            		break;
            	case "TITLE3":
            		row = sheet.getRow(15);
            		cell = row.getCell(3);
            		cell.setCellValue(entry.getValue());
            		break;
            	case "IE1":
            		row = sheet.getRow(17);
            		cell = row.getCell(5);
            		cell.setCellValue(entry.getValue());
            		break;
            	case "IE1FOR":
            		row = sheet.getRow(17);
            		cell = row.getCell(3);
            		cell.setCellValue(entry.getValue());
            		break;
            	case "IE2":
            		row = sheet.getRow(18);
            		cell = row.getCell(5);
            		cell.setCellValue(entry.getValue());
            		break;
            	case "IE2FOR":
            		row = sheet.getRow(18);
            		cell = row.getCell(3);
            		cell.setCellValue(entry.getValue());
            		break;
            	case "IE3":
            		row = sheet.getRow(19);
            		cell = row.getCell(5);
            		cell.setCellValue(entry.getValue());
            		break;
            	case "IE3FOR":
            		row = sheet.getRow(19);
            		cell = row.getCell(3);
            		cell.setCellValue(entry.getValue());
            		break;
            	case "TTM":
            		row = sheet.getRow(23);
            		cell = row.getCell(5);
            		cell.setCellValue(entry.getValue());
            		break;
            	case "TTE":
            		row = sheet.getRow(25);
            		cell = row.getCell(5);
            		cell.setCellValue(entry.getValue());
            		break;
            	}
            }

//	            for (int i = firstLine; i <= lines; i++) {
//	                XSSFRow row = sheet.getRow((short) i);
//	                if (null == row) {
//	                    continue;
//	                } else {
//	                    Cell cell1 = row.getCell((short) 5);
//	                    System.out.println("row.getFirstCellNum() = " + row.getFirstCellNum());
//	                    Cell cell = row.getCell(row.getFirstCellNum());
//	                    
//	                    if (null == cell) {
//	                        continue;
//	                    } else {
//	                        System.out.println(cell.getNumericCellValue());
//	                        int temp = (int) cell.getNumericCellValue();
//	                        cell.setCellValue(temp + 1);
//	                    }
//	                }
//	            }
            
            
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