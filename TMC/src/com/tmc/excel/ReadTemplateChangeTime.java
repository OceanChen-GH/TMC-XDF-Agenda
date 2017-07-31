package com.tmc.excel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;  
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook; 

public class ReadTemplateChangeTime {
	
	/* map:����
	{Word of the day=Thanksgiving, TIMER(agenda)=Gen&Victor, Ah Counter=Ocean, Warm up Session=Ocean, SAA=Tracy, TOM(review)=Alwin, IE1=Harry, IE3=Ivy Chen��CL1��, Greeter=Ivy, IE2=Carol Fang , GRAMMARIAN=Melody, SPEAKER2=Tracy(CC1), SPEAKER1=Elena(CC5), SPEAKER3=Harry(CC7), GE=Jamie}
	
	ע�⣺map��TIMER(agenda) change to TIMER; TOM(review) change to TOM
	{Word of the day=merry, IE1For=IE For Angelia , TTE=Elena, IE3For=IE For Tracy , IE2For=IE For Lancy , TIMER=Ivy, Ah Counter=Paul, TTM=Angelia, Title1= Colorful Life Makes an Outgoing Girl, Title3= Hey��it's me, Title2= Story of My Name, Warm up Session=Ivy, SAA=Judy, TOM=Jason, IE1=Tina, IE3=, Greeter=Elena, IE2=, Theme=Christmas Eve, Ice Breaking Eve, GRAMMARIAN=Tina, SPEAKER2=Lancy (CC1), SPEAKER1=Angelia (CC1), SPEAKER3=Tracy (CC1), GE=Alwin}
	*/
	
	public static void main(String[] args) throws IOException {
		// put roles into rolesMap
		//readF1("E:/TMC-XDF-Agenda-Tools/roles-40th.txt");
		// modify Excel according to rolesMap
//		changeCell("E:/@Ӣ��ѧϰ/TMC/@Agenda-Tools/AgendaGenerator-XDF/agenda-xxth.xlsx");//����ʹ��
		changeCell("./agenda-xxth.xlsx");//���ʹ��
	}
	
    public static final void readF1(String filePath) throws IOException {}  
  
    public static final void readF2(String filePath) throws IOException {}
    
    public static final void changeCell(String excelFile){
        String fileToBeRead = excelFile;// "E:/test.xlsx"; // excelλ��
        try {
        	SimpleDateFormat df=new SimpleDateFormat("HH:mm");
        	
            XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(
                    fileToBeRead));
            XSSFSheet sheet = workbook.getSheet("Sheet1");
            int linesNum = sheet.getLastRowNum();//��ȡ���һ�е�����
            XSSFRow row = null;
            XSSFRow nextRow = null;
            Cell cellCostTime = null;// ȡ��һ���ڵ���ʱ
            Cell cellPreviousClockTime = null;// ��д��ʱ�䵥Ԫ��
        	Cell cellClockTime = null;// ��д��ʱ�䵥Ԫ��
        	   
            for (int i = 3; i < (linesNum - 10); i++) {// ��������10��
            	row = sheet.getRow(i);
            	cellPreviousClockTime = row.getCell(2);
            	String startTimeStr = null;
            	if (i == 3) {
//            		startTimeStr = String.valueOf(cellPreviousClockTime.getDateCellValue());
            		startTimeStr = df.format(cellPreviousClockTime.getDateCellValue());
            	} else {
            		startTimeStr = cellPreviousClockTime.getStringCellValue();
            	}
            	Date startTime = null;
//            	startTimeStr = df.format(cellPreviousClockTime.getDateCellValue());
            	try {
					startTime = df.parse(startTimeStr);
				} catch (ParseException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
            	Calendar cal=Calendar.getInstance();
            	cal.setTime(startTime);
            	
        		cellCostTime = row.getCell(4);// E�У���5�У���0����
        		String costTimeStr = cellCostTime.getStringCellValue();
        		if (!costTimeStr.equals("")) {
	        		int costTime = 0;
	        		if (costTimeStr.contains("-")) {
	        			costTime = Integer.parseInt(costTimeStr.substring(costTimeStr.indexOf("-") + 1, costTimeStr.length()-1));
	        		} else {
	        			costTime = Integer.parseInt(costTimeStr.substring(0, costTimeStr.length()-1));
	        		}
	            	cal.add(Calendar.MINUTE, +costTime);
	            	System.out.println(df.format(cal.getTime()));

	            	nextRow = sheet.getRow(i + 1);
	            	if (!nextRow.getCell(4).getStringCellValue().equals("")) {
//	        			row = sheet.getRow(i + 1);// д��ĵ�Ԫ��������һ��
	        			cellClockTime = nextRow.getCell(2);
	        			cellClockTime.setCellValue(df.format(cal.getTime()));
	            	} else {
	            		i++;
	            		nextRow = sheet.getRow(i + 1);
	            		cellClockTime = nextRow.getCell(2);
	        			cellClockTime.setCellValue(df.format(cal.getTime()));
	            	}
        		} else {
        			i++;
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