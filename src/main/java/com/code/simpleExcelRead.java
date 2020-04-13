/*package com.code;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import org.apache.commons.compress.archivers.dump.InvalidFormatException;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.testng.annotations.Test;


public class simpleExcelRead 
{

	 @Test
	  public void excel() throws EncryptedDocumentException, InvalidFormatException, FileNotFoundException, IOException { 
	                //get the excel sheet file location 
		 InputStream inp = new FileInputStream("d:xceltask.xlsx");
		 Workbook wb = WorkbookFactory.create(inp);
	                
			            Sheet sh = wb.getSheet("sheet1");
	                //get the total row count in the excel sheet	
			            int rowcount = sh.getLastRowNum();
			              for (int i = 0; i <= rowcount; i++) 
	                  {
	                    // get the total cell count in the excel sheet
				               int cellcount = sh.getRow(i).getLastCellNum();
				                  for (int j = 0; j < cellcount; j++) 
	                        {                         
	                            //get cell value at the given position [i][j]
			                        String value = sh.getRow(i).getCell(j).getStringCellValue();
	                            //print the cell value
					                    System.out.println(value);
					
				                   }			
			                }	
			          }
}*/