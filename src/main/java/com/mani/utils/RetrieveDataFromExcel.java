package com.mani.utils;
import java.io.File;
import java.io.FileInputStream;
import java.util.HashMap;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class RetrieveDataFromExcel {
	public static HashMap<String, String> metaTagsData;
	public static String title;
	public static HashMap<String, String> readfromExcel(String TCName) throws Exception{
		metaTagsData = new HashMap<>();
		DataFormatter formtter = new DataFormatter();
		File file = new File("./TestDataFile.xlsx");
		FileInputStream fis = new FileInputStream(file);
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		XSSFSheet sheet = wb.getSheetAt(0);
		
		int noOfRows = sheet.getLastRowNum();
		for(int i=0;i<noOfRows;i++) {
			String validateTitle = formtter.formatCellValue(sheet.getRow(i).getCell(1));
			if(validateTitle.trim().equalsIgnoreCase(TCName)) {
				int columns = sheet.getRow(i).getLastCellNum();
				for (int j=0;j<columns;j++) {
					String getFieldvalue = formtter.formatCellValue(sheet.getRow(i).getCell(j)); 
					if(getFieldvalue!="") {
						System.out.println(formtter.formatCellValue(sheet.getRow(0).getCell(j)) +" :"+getFieldvalue);
					}
					String colHeader=null;
					colHeader = sheet.getRow(0).getCell(j).getStringCellValue();
					metaTagsData.put(colHeader, getFieldvalue);
				}
			}
		}		
		return metaTagsData;
	}

}
