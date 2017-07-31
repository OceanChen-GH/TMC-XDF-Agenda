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

public class ReadTemplateForXZHSH {
	
	/* map:例子
	{Word of the day=Thanksgiving, TIMER(agenda)=Gen&Victor, Ah Counter=Ocean, Warm up Session=Ocean, SAA=Tracy, TOM(review)=Alwin, IE1=Harry, IE3=Ivy Chen（CL1）, Greeter=Ivy, IE2=Carol Fang , GRAMMARIAN=Melody, SPEAKER2=Tracy(CC1), SPEAKER1=Elena(CC5), SPEAKER3=Harry(CC7), GE=Jamie}
	
	注意：map中TIMER(agenda) change to TIMER; TOM(review) change to TOM
	{Word of the day=merry, IE1For=IE For Angelia , TTE=Elena, IE3For=IE For Tracy , IE2For=IE For Lancy , TIMER=Ivy, Ah Counter=Paul, TTM=Angelia, Title1= Colorful Life Makes an Outgoing Girl, Title3= Hey，it's me, Title2= Story of My Name, Warm up Session=Ivy, SAA=Judy, TOM=Jason, IE1=Tina, IE3=, Greeter=Elena, IE2=, Theme=Christmas Eve, Ice Breaking Eve, GRAMMARIAN=Tina, SPEAKER2=Lancy (CC1), SPEAKER1=Angelia (CC1), SPEAKER3=Tracy (CC1), GE=Alwin}
	*/
	private static Map<String, String> rolesMap = new HashMap<String, String>();
	
	public static void main(String[] args) throws IOException {
		// put roles into rolesMap
		readF1("./角色报名表-心之声头马国际演讲俱乐部.txt");
//		readF1(args[0]);
		// modify Excel according to rolesMap
		changeCell("./会议流程(Agenda).xlsx");
//		changeCell(args[1]);
	}
	
    public static final void readF1(String filePath) throws IOException {  
        BufferedReader br = new BufferedReader(new InputStreamReader(  
                new FileInputStream(filePath), "UTF-8"));  
        String greeter = null;
        for (String line = br.readLine(); line != null; line = br.readLine()) {  
            if (line.length() >= 5 ) {
            	if (line.substring(0, 4).equalsIgnoreCase("迎宾官1")) {
//            		rolesMap.put(line.substring(0, 4),	line.substring(6));
            		greeter = line.substring(6);
            		continue;
            	}
            	if (line.substring(0, 4).equalsIgnoreCase("迎宾官2")) {
            		if (line.length() > 6) {
            			greeter = greeter + "，" + line.substring(6);// 两个迎宾官合并
            		}
            		rolesMap.put(line.substring(0, 3),	greeter);
            		continue;
            	}
            	if (line.substring(0, 4).equalsIgnoreCase("即兴主持")) {
            		rolesMap.put(line.substring(0, 4),	line.substring(6));
            		continue;
            	}
            	if (line.substring(0, 4).equalsIgnoreCase("即兴点评")) {
            		rolesMap.put(line.substring(0, 4),	line.substring(6));
            		continue;
            	}
            	if (line.substring(0, 3).equalsIgnoreCase("安保官")) {
            		rolesMap.put(line.substring(0, 3),	line.substring(5));
            		continue;
            	}
            	if (line.substring(0, 3).equalsIgnoreCase("主持人")) {
            		rolesMap.put(line.substring(0, 3),	line.substring(5));
            		continue;
            	}
            	if (line.substring(0, 3).equalsIgnoreCase("总点评")) {
            		rolesMap.put(line.substring(0, 3),	line.substring(5));
            		continue;
            	}
            	if (line.substring(0, 3).equalsIgnoreCase("时间官")) {
            		rolesMap.put(line.substring(0, 3),	line.substring(5));
            		continue;
            	}
            	if (line.substring(0, 3).equalsIgnoreCase("赘语官")) {
            		rolesMap.put(line.substring(0, 3),	line.substring(5));
            		continue;
            	}
            	if (line.substring(0, 3).equalsIgnoreCase("文法官")) {
            		rolesMap.put(line.substring(0, 3),	line.substring(5));
            		continue;
            	}
            	if (line.substring(0, 3).equalsIgnoreCase("备稿1")) {
            		rolesMap.put(line.substring(0, 3),	line.substring(5));
            		continue;
            	}
            	if (line.substring(0, 3).equalsIgnoreCase("点评1")) {
            		rolesMap.put(line.substring(0, 3),	line.substring(5));
            		continue;
            	}
            	if (line.substring(0, 3).equalsIgnoreCase("备稿2")) {
            		rolesMap.put(line.substring(0, 3),	line.substring(5));
            		continue;
            	}
            	if (line.substring(0, 3).equalsIgnoreCase("点评2")) {
            		rolesMap.put(line.substring(0, 3),	line.substring(5));
            		continue;
            	}
            	if (line.substring(0, 3).equalsIgnoreCase("备稿3")) {
            		rolesMap.put(line.substring(0, 3),	line.substring(5));
            		continue;
            	}
            	if (line.substring(0, 3).equalsIgnoreCase("点评3")) {
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
        String fileToBeRead = excelFile; // excel位置
        try {
            XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(
                    fileToBeRead));
            XSSFSheet sheet = workbook.getSheet("Agenda");
//            int firstLine = 3;// role从第4行开始
//            int lines = sheet.getLastRowNum();
            // 变量rolesMap中的元素
            for (Map.Entry<String, String> entry : rolesMap.entrySet()) {
            	XSSFRow row = null;
            	Cell cell = null;
            	Cell cell2 = null;// Timer等可能需要修改两行
            	switch (entry.getKey()) {
            	case "迎宾官":
            		row = sheet.getRow(6);
            		cell = row.getCell(7);
            		cell.setCellValue(entry.getValue());
            		
            		row = sheet.getRow(15);
            		cell = row.getCell(7);
            		cell.setCellValue(entry.getValue());
            		
            		row = sheet.getRow(32);
            		cell = row.getCell(7);
            		cell.setCellValue(entry.getValue());
            		break;
            	case "安保官":
            		row = sheet.getRow(7);
            		cell = row.getCell(7);
            		cell.setCellValue(entry.getValue());
            		break;
            	case "主持人":
            		row = sheet.getRow(8);
            		cell = row.getCell(7);
            		cell.setCellValue(entry.getValue());
            		
            		row = sheet.getRow(10);
            		cell2 = row.getCell(7);
            		cell2.setCellValue(entry.getValue());
            		
            		row = sheet.getRow(30);// 活动评价
            		cell = row.getCell(7);
            		cell.setCellValue(entry.getValue());
            		break;
            	case "总点评":
            		row = sheet.getRow(11);
            		cell = row.getCell(7);
            		cell.setCellValue(entry.getValue());
            		
            		row = sheet.getRow(29);
            		cell2 = row.getCell(7);
            		cell2.setCellValue(entry.getValue());
            		break;
            	case "时间官":
            		row = sheet.getRow(12);
            		cell = row.getCell(7);
            		cell.setCellValue(entry.getValue());
            		
            		row = sheet.getRow(28);
            		cell = row.getCell(7);
            		cell.setCellValue(entry.getValue());
            		break;
            	case "赘语官":
            		row = sheet.getRow(13);
            		cell = row.getCell(7);
            		cell.setCellValue(entry.getValue());
            		
            		row = sheet.getRow(25);
            		cell = row.getCell(7);
            		cell.setCellValue(entry.getValue());
            		break;
            	case "文法官":
            		row = sheet.getRow(14);
            		cell = row.getCell(7);
            		cell.setCellValue(entry.getValue());
            		
            		row = sheet.getRow(26);
            		cell = row.getCell(7);
            		cell.setCellValue(entry.getValue());
            		break;
            	case "即兴主持":
            		row = sheet.getRow(16);
            		cell = row.getCell(7);
            		cell.setCellValue(entry.getValue());
            		break;
            	case "即兴点评":
            		row = sheet.getRow(17);
            		cell = row.getCell(7);
            		cell.setCellValue(entry.getValue());
            		break;
            	case "备稿1":
            		row = sheet.getRow(19);
            		cell = row.getCell(7);
            		cell.setCellValue(entry.getValue());
            		break;
            	case "备稿2":
            		row = sheet.getRow(20);
            		cell = row.getCell(7);
            		cell.setCellValue(entry.getValue());
            		break;
            	case "备稿3":
            		row = sheet.getRow(21);
            		cell = row.getCell(7);
            		cell.setCellValue(entry.getValue());
            		break;
            	case "点评1":
            		row = sheet.getRow(22);
            		cell = row.getCell(7);
            		cell.setCellValue(entry.getValue());
            		break;
            	case "点评2":
            		row = sheet.getRow(23);
            		cell = row.getCell(7);
            		cell.setCellValue(entry.getValue());
            		break;
            	case "点评3":
            		row = sheet.getRow(24);
            		cell = row.getCell(7);
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