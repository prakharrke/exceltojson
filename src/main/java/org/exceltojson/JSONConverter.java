package org.exceltojson;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Set;

import org.json.JSONArray;
import org.json.JSONObject;

import com.aspose.cells.Cell;
import com.aspose.cells.Cells;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.google.gson.Gson;

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
			
			excel.close();
			workbook.save(path);
			
			
		}else {
			
			System.out.println("File does not exist");
			return;
		}
	}
	
	public JSONObject convertToJSONRowWise(String path, String sheetName) throws Exception {
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
				row.put("edited",false);
				row.put("rowNum",i);
				excelJson.put("row" + i, row);
				
			}
			
			System.out.println(excelJson);
			excel.close();
			workbook.save(path);
		
			return excelJson;
		}
		
		return new JSONObject();
	}
	
	public void updateExcel(String path,String sheetName,JSONObject updatedJSON) throws Exception {
		excelJson = new JSONObject();
		File excelFile = new File(path);
		if(excelFile.exists()) {
			FileInputStream excel = new FileInputStream(excelFile);
			Workbook workbook = new Workbook(excel);
			Worksheet sheet = workbook.getWorksheets().get(sheetName);
			Cells cells = sheet.getCells();
			columnCount = cells.getMaxDataColumn();
			excelJson = updatedJSON;
			Object[] keys = excelJson.keySet().toArray();
			try {
			for(Object rowKey : keys) {
				
				if(excelJson.getJSONObject(rowKey.toString()).getBoolean("edited")) {
					JSONObject rowObject = new JSONObject();
					rowObject = excelJson.getJSONObject(rowKey.toString());
					int rowNum = rowObject.getInt("rowNum");
					Object[] headers = rowObject.keySet().toArray();
					Set<String>headersSet = rowObject.keySet();
					int colIndex = 0;
					for(int i=0;i<=(columnCount+2);i++) {
						if(headers[i].toString().equalsIgnoreCase("edited") ||  headers[i].toString().equalsIgnoreCase("rowNum")) {
							continue;
						}else {
							// CALCULATE THE COLUMN INDEX WITH HEADER SAME AS headers[i]
						for(int j=0;j<=columnCount;j++) {
							Cell c = cells.get(0,j);
							if(c.getValue().equals(headers[i].toString())) {
								colIndex = j;
								break;
							}
						}
						Cell cell = cells.get(rowNum,colIndex);
						cell.putValue(rowObject.get(headers[i].toString()));
						colIndex++;
						}
					}
					
					
					
					// UPDATE THE EXCEL
				}else {
					continue;
				}
			}
			excel.close();
			workbook.save(path);
			}catch(Exception e) {
				e.printStackTrace();
				excel.close();
				workbook.save(path);
			}
		}
	}

	
	
	
	
	
	
	
	
}
