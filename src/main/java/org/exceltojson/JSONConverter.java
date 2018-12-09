package org.exceltojson;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.ArrayList;
import java.util.List;

import org.json.JSONArray;
import org.json.JSONObject;

import com.aspose.cells.Cell;
import com.aspose.cells.Cells;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

public class JSONConverter {
	private int lastRowCount=0;
	private int columnCount=0;
	List<String>columns = new ArrayList<String>();
	JSONObject excelJson = new JSONObject();
	
	public void  convertToJSONColumnWise(String path, String sheetName) throws Exception {
		excelJson = new JSONObject();
		File excelFile = new File(path);
		if(excelFile.exists()) {
			FileInputStream excel = new FileInputStream(excelFile);
			Workbook workbook = new Workbook(excel);
			Worksheet sheet= workbook.getWorksheets().get(sheetName);
			Cells cells = sheet.getCells();
			 lastRowCount = cells.getMaxRow();
			columnCount = cells.getMaxColumn();
			
			for(int i=0; i<=columnCount; i++) {
				Cell cell = cells.get(0,i);
				columns.add(cell.getValue().toString());
				
			}
			
			for(String columnName : columns) {
				JSONArray columnValues = new JSONArray();
				int columnNumber = columns.indexOf(columnName);
				
				for(int i=1;i<=lastRowCount;i++) {
					Cell cellValue = cells.get(i,columnNumber);
					cellValue.putValue("Changed");
					columnValues.put(cellValue.getValue());
				}
				
				excelJson.put(columnName,columnValues);
				
			}
			
			System.out.println(excelJson);
			
			
			workbook.save(path);
			excel.close();
			
		}else {
			
			System.out.println("File does not exist");
			return;
		}
	}
	
	public void convertToJSONRowWise(String path, String sheetName) throws Exception {
		excelJson = new JSONObject();
		columns = new ArrayList<String>();
		File excelFile = new File(path);
		if(excelFile.exists()) {
			FileInputStream excel = new FileInputStream(excelFile);
			Workbook workbook = new Workbook(excel);
			Worksheet sheet = workbook.getWorksheets().get(sheetName);
			Cells cells = sheet.getCells();
			lastRowCount = cells.getMaxDataRow();
			columnCount = cells.getMaxDataColumn();
			for(int i=0;i<=columnCount;i++) {
				Cell cell = cells.get(0,i);
				columns.add(cell.getValue().toString());
			}
			for(int i=1; i<=lastRowCount; i++) {
				JSONObject row = new JSONObject();
				for(String columnName : columns) {
					Cell cell = cells.get(i,columns.indexOf(columnName));
					row.put(columnName, cell.getValue());
				}
				
				excelJson.put("row" + i, row);
				
			}
			
			System.out.println(excelJson);
		}
		
		
	}

}
