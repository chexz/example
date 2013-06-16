package ck;

import java.io.*;

import jxl.*;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import ck.*;

public class EntryApp {// an application

	
		public static void main(String args[]) {

			//System.out.println(System.getProperty("user.dir"));
			
			//save input file name
			String[] fileName = new String[args.length];
			
			for(int i=0;i<args.length;i++)
			{ 
				//System.out.println(args[i]);
				fileName[i] = args[i];
				//System.out.println(fileName[i]);
			}	

			//open elementTableFile and save element to ElementTable array
			Workbook elementBook = null;
			try
			{	
				elementBook = Workbook.getWorkbook(new File(fileName[0]));
			}
		    catch (Exception e)   
		    {  
		    	System.out.println(e);  
		    }   
	
		     
			if(elementBook==null)
			{
				System.out.println("Get workbook failed!");
				return;
			}
			
			Sheet sheet = elementBook.getSheet(0); //get the fist sheet
			
			int rowNum = sheet.getRows();
			//System.out.println(rowNum); 
			
			int columnNum = sheet.getColumns();
			//System.out.println(columnNum);
			
			ElementTable[] element = new ElementTable[rowNum-1];
			
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
				element[index] = new ElementTable();
				element[index].elementName = sheet.getCell(0, i).getContents();
				element[index].xMax = Double.parseDouble(sheet.getCell(xMaxColumn, i).getContents());
				element[index].xMin = Double.parseDouble(sheet.getCell(xMinColumn, i).getContents());
				element[index].yMax = Double.parseDouble(sheet.getCell(yMaxColumn, i).getContents());
				element[index].yMin = Double.parseDouble(sheet.getCell(yMinColumn, i).getContents());
				//System.out.println("element["+index+"].elementName="+element[index].elementName+"element["+index+"].xMax="+element[index].xMax); 
				//System.out.println("element["+index+"].yMin="+element[index].yMin+"element["+index+"].yMax="+element[index].yMax);  
			}
			
			elementBook.close();
			
			//open sample file, read all data and record all data 
			String[] sampleName = new String[fileName.length-2];
			String[][] value = new String[element.length][sampleName.length];
			
			for(int i=1; i<fileName.length-1;i++)
			{
				String tempS[] = fileName[i].split("\\.");
				sampleName[i-1] = tempS[0];
				//System.out.println("tempS[0]="+tempS[0]+"  sampleName["+i+"]="+sampleName[i-1]);
				
				try
				{
					Workbook book = Workbook.getWorkbook(new File(fileName[i]));
					sheet = book.getSheet(0);
					rowNum = sheet.getRows();
					columnNum = sheet.getColumns();
					
					int xColumn = -1;
					int yColumn = -1;
					int valueColumn = -1;
					
					for(int j=0; j<columnNum; j++)
					{
						String cellCnt = sheet.getCell(j, 0).getContents();
						//System.out.println(cellCnt);
						if(cellCnt.compareToIgnoreCase("x")==0)
						{
							xColumn = j;
						}
						else if(cellCnt.compareToIgnoreCase("y")==0)
						{
							yColumn = j;
						}
						else if(cellCnt.compareToIgnoreCase("value")==0)
						{
							valueColumn = j;
						}
						
						//System.out.println("x="+xColumn+"y="+yColumn+"value="+valueColumn);  
					}
					
					for(int j=1; j<rowNum; j++)
					{
						double x =  Double.parseDouble(sheet.getCell(xColumn, j).getContents());
						double y =  Double.parseDouble(sheet.getCell(yColumn, j).getContents());
						
						for(int k=0; k<element.length; k++)
						{
							ElementTable eTemp  = element[k];
							if((x>=eTemp.xMin)&&(x<=eTemp.xMax)&&(y>=eTemp.yMin)&&(y<=eTemp.yMax))
							{
								value[k][i-1] = sheet.getCell(valueColumn, j).getContents();
							}
						}
						
					}
					book.close();
				}
				catch (Exception e)   
			    {  
			    	System.out.println(e);  
			    } 
			}

			//sample data write to new excel file
			try
			{
				WritableWorkbook book = Workbook.createWorkbook(new File(fileName[fileName.length-1]));
				WritableSheet wrtSheet = book.createSheet("the first page", 0);
				
				for(int i=0; i<element.length;i++)
				{
					Label label=new Label(0,i+1,element[i].elementName);
					wrtSheet.addCell(label);
				}
				
				for(int i=0; i<sampleName.length;i++)
				{
					Label label=new Label(i+1,0,sampleName[i]);
					wrtSheet.addCell(label);
				}
				
				for(int i=0; i<element.length;i++)
				{
					for(int j=0; j<sampleName.length;j++)
					{
						Label label=new Label(j+1,i+1,value[i][j]);
						wrtSheet.addCell(label);
					}
				}
				book.write();
				book.close();
								
			}
			catch (Exception e) 
			{
				System.out.println(e);
			}
		}

}
