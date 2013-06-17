package ck;

import java.io.File;
import jxl.Sheet;
import jxl.Workbook;

public class ElementTable {
	static int ERROR_CODE_OPEN_FILE_FAILED = -1;
	ElementProperty[] elements;
	public void loadElementTable(String filename) {
		Workbook elementBook = null;
		try {	
			elementBook = Workbook.getWorkbook(new File(filename));
		} catch (Exception e) {  
	    	System.err.println(e);  
	    }   
	     
		if(elementBook == null) {
			System.err.println("load element table file failed!");
			System.exit(ERROR_CODE_OPEN_FILE_FAILED);
			return;
		}
		
		// get the fist sheet. the other sheet is ignored
		Sheet sheet = elementBook.getSheet(0); 
		
		int rowNum = sheet.getRows();
		System.out.println("" + (rowNum - 1) + "elements in the table."); 
		
		int columnNum = sheet.getColumns();
		
		ElementProperty[] element = new ElementProperty[rowNum-1];
		int xMaxColumn = -1;
		int xMinColumn = -1;
		int yMaxColumn = -1;
		int yMinColumn = -1;
		
		for(int i=1; i<columnNum; i++)
		{
			String cellCnt = sheet.getCell(i, 0).getContents();
			//System.out.println(cellCnt);
			if(cellCnt.compareToIgnoreCase("xMax")==0)
			{
				xMaxColumn = i;
			}
			else if(cellCnt.compareToIgnoreCase("xMin")==0)
			{
				xMinColumn = i;
			}
			else if(cellCnt.compareToIgnoreCase("yMax")==0)
			{
				yMaxColumn = i;
			}
			else if(cellCnt.compareToIgnoreCase("yMin")==0)
			{
				yMinColumn = i;
			}
			
			//System.out.println("xMax="+xMaxColumn+"xMin="+xMinColumn+"yMax="+yMaxColumn+"yMin="+yMinColumn);  
		}
		
		for(int i=1; i<rowNum; i++)
		{
			int index = i-1;
			element[index] = new ElementProperty();
			element[index].elementName = sheet.getCell(0, i).getContents();
			element[index].xMax = Double.parseDouble(sheet.getCell(xMaxColumn, i).getContents());
			element[index].xMin = Double.parseDouble(sheet.getCell(xMinColumn, i).getContents());
			element[index].yMax = Double.parseDouble(sheet.getCell(yMaxColumn, i).getContents());
			element[index].yMin = Double.parseDouble(sheet.getCell(yMinColumn, i).getContents());
			//System.out.println("element["+index+"].elementName="+element[index].elementName+"element["+index+"].xMax="+element[index].xMax); 
			//System.out.println("element["+index+"].yMin="+element[index].yMin+"element["+index+"].yMax="+element[index].yMax);  
		}
		
		elementBook.close();
		
		elements = element;
	}
	
	String classifyElements(double x, double y) {
		for (int i = 0; i < elements.length; i++) {
			if (elements[i].xMin <= x && elements[i].xMax >= x 
					&& elements[i].yMin <= y && elements[i].yMax >= y) {
				return elements[i].elementName;
			}
		}
		return "";
	}
}
