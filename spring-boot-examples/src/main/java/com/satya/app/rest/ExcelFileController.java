package com.satya.app.rest;

import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import com.satya.app.model.TestModel;
import com.satya.app.util.ExcelToPojoUtils;

@RestController
public class ExcelFileController {
	
	private static final String FILE_LOCATION = "C:\\Users\\satya\\Documents\\Book1.xlsx";
	
	public void readExcelFile() {
		try(FileInputStream file = new FileInputStream(new File(FILE_LOCATION))) {
            // Create Workbook instance holding reference to .xlsx file 
            XSSFWorkbook workbook = new XSSFWorkbook(file); 
  
            // Get first/desired sheet from the workbook 
            XSSFSheet sheet = workbook.getSheetAt(0); 
  
            // Iterate through each rows one by one 
            Iterator<Row> rowIterator = sheet.iterator();
            
            List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();
            
            while (rowIterator.hasNext()) { 
            	Row row = rowIterator.next(); 
            	  
                // For each row, iterate through all the columns 
                Iterator<Cell> cellIterator = row.cellIterator();
                while (cellIterator.hasNext()) {
                	Cell cell = cellIterator.next();
                }
            }
            
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	@GetMapping("/readExcel")
	public String readExcel() {
		try(FileInputStream file = new FileInputStream(new File(FILE_LOCATION))) {
			  List<TestModel> list = ExcelToPojoUtils.toPojoList(TestModel.class, file);
			  
			  for(TestModel test : list) {
				  System.out.println(test.getHeader1());
				  System.out.println(test.getuS()+" "+test.getpT()+" "+test.getcH());
			  }
            
		} catch (Exception e) {
			e.printStackTrace();
		}
		return "Done";
	}

}
