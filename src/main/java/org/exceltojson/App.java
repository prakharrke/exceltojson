package org.exceltojson;

/**
 * Hello world!
 *
 */
public class App 
{
    public static void main( String[] args ) throws Exception
    {
    	JSONConverter conv = new JSONConverter();
    	conv.convertToJSONRowWise("C:\\Users\\Prakhar\\Documents\\Book1.xlsx","Sheet1");
    }
}
