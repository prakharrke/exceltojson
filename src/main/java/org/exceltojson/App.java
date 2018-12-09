package org.exceltojson;

import org.json.JSONObject;

/**
 * Hello world!
 *
 */
public class App 
{
    public static void main( String[] args ) throws Exception
    {
    	JSONConverter conv = new JSONConverter();
    	JSONObject resp = conv.convertToJSONRowWise("C:\\Users\\Prakhar\\Documents\\Book1.xlsx","Sheet1");
    	JSONObject temp = new JSONObject ();
    	temp = resp.getJSONObject("row2");
    	temp.put("Header1","Updated1");
    	temp.put("Header2","Updated2");
    	temp.put("Header3","Updated3");
    	temp.put("Header4","Updated4");
    	temp.put("Header5","Updated5");
    	temp.put("Header6","Updated6");
    	temp.put("edited",true);
    	resp.put("row2",temp);
    	conv.updateExcel("C:\\Users\\Prakhar\\Documents\\Book1.xlsx","Sheet1",resp );
    	//conv.convertToJSONColumnWise("C:\\Users\\Prakhar\\Documents\\Book1.xlsx", "Sheet1");
    }
}
