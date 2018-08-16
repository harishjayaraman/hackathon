package com.hackathon.coreloop;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;



public class ExcelReader {

	 public static final String SAMPLE_XLSX_FILE_PATH = "C:/Users/harish.jayaraman/Desktop/coreloop.xlsx";
	 
	 
	 public static Map<Object, String> getTestDetails(String filePath) throws IOException{
		 
	        FileInputStream inputStream = new FileInputStream(new File(SAMPLE_XLSX_FILE_PATH));
	       Map<Object, String> classAndTest = new HashMap<Object, String>();
	        Workbook workbook = new XSSFWorkbook(inputStream);
	        Sheet firstSheet = workbook.getSheetAt(0);
	        Iterator<Row> rowIterator = firstSheet.iterator();
	         
	        while (rowIterator.hasNext()) {
	            Row row = rowIterator.next();
	            Iterator<Cell> cellIterator = row.cellIterator();
	            Object testClass = null;
            	String methodName = null;
	            while(cellIterator.hasNext())
	            {
	            	Cell cell = cellIterator.next();
	            	if(cell.getColumnIndex() == 2 && cell.getRowIndex() >0)
	            	{
	            		String className = cell.getStringCellValue();
	            		try {
	            			testClass = Class.forName(className).newInstance();
							
						} catch (InstantiationException e) {
							// TODO Auto-generated catch block
							e.printStackTrace();
						} catch (IllegalAccessException e) {
							// TODO Auto-generated catch block
							e.printStackTrace();
						} catch (ClassNotFoundException e) {
							// TODO Auto-generated catch block
							e.printStackTrace();
						}
	            	}
	            	if(cell.getColumnIndex() == 3 &&  cell.getRowIndex() >0)
	            	{
	            		methodName = cell.getStringCellValue();
	            		classAndTest.put(testClass, methodName);
	            	}
	            	
	            }
	        }
	        workbook.close();
	        inputStream.close();
			return classAndTest;
	    }

	 
	 public static void main(String...strings)
	 {
		 try {
			 Map<Object, String>  classAndTest = getTestDetails(SAMPLE_XLSX_FILE_PATH);
			 
			 for(Map.Entry<Object, String> ct : classAndTest.entrySet()){
				 Class clazz = ct.getKey().getClass();
				 Object t = null;
				 try {
					 t = clazz.newInstance();
				} catch (InstantiationException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				} catch (IllegalAccessException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				}
				 try {
					Method method = clazz.getDeclaredMethod(ct.getValue(), null);
					try {
						method.invoke(t, null);
					} catch (IllegalAccessException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					} catch (IllegalArgumentException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					} catch (InvocationTargetException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
				} catch (NoSuchMethodException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				} catch (SecurityException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			 }
			 
			 
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	 }
	        
	
}
