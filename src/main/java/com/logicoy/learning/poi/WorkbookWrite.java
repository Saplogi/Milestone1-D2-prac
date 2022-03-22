package com.logicoy.learning.poi;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.log4j.BasicConfigurator;
import org.apache.log4j.Logger;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.formula.functions.NumericFunction;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WorkbookWrite {
	
	static Logger log = Logger.getLogger("WorkbookWrite");
	
	public static void main(String[] args){
		BasicConfigurator.configure();
		String p ="C:\\Users\\SaptarshiChakraborty\\Documents\\Training\\Day.xlsx";
		
		try {
			FileInputStream fis = new FileInputStream(p);
			XSSFWorkbook wb = new XSSFWorkbook(fis);
			
			XSSFSheet s = wb.getSheet("Sheet1");
			//HSSFSheet s = wb.getSheet("Sheet1");
			
			Iterator<Row> it =s.iterator();
			
			while(it.hasNext()) {
				XSSFRow row = (XSSFRow) it.next();
				Iterator<Cell> itc = row.cellIterator();
				while (itc.hasNext())
	            {
	                Cell cell = (Cell) itc.next();
	                //Check the cell type and format accordingly
	                switch (cell.getCellType()) 
	                {
	                    case NUMERIC:
	                       // log.debug(cell.getNumericCellValue());
	                    	log.info(cell.getNumericCellValue()+"     ");
	                    	//System.out.print(cell.getNumericCellValue()+"     ");
	                        break;
	                    case STRING:
	                        //log.debug(cell.getStringCellValue());
	                    	log.info(cell.getStringCellValue()+"     ");
	                        break;
					default:
						break;
	                }
	            }
	            System.out.println("");
	        }
			
			wb.close();
			fis.close();
		}
		catch(FileNotFoundException e) {
			log.error("File Not Found Type error"+" "+e.getMessage());
		}
		catch(IOException e) {
			log.error("IO Exception"+" "+e.getMessage());
		}
		
		
		//HSSFWorkbook wb = new HSSFWorkbook(fis);
		
		
		
	}
}
