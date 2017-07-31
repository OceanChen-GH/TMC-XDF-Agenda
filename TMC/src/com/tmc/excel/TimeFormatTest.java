package com.tmc.excel;

import java.text.SimpleDateFormat;
import java.util.Date;

public class TimeFormatTest {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		Date date=new Date();
		System.out.println(date);
		SimpleDateFormat sdf=new SimpleDateFormat("yyyy-MM-dd");
		String sDate=sdf.format(date);
		System.out.println(sDate);
	}

}
