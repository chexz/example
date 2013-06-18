package ck;

import java.io.*;
import java.lang.Thread.UncaughtExceptionHandler;
import java.util.Vector;

import jxl.*;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

public class EntryApp implements UncaughtExceptionHandler {
	String elementTableFilename = "elementTable.xls";
	String resultFilename = "result.xls";
	Vector<String> sourceFilenames = new Vector<String>();
	
	public String getElementTableFilename() {
		return elementTableFilename;
	}
	
	public String getResultFilename() {
		return resultFilename;
	}
	
	public Vector<String> getSourceFilenames() {
		return sourceFilenames;
	}

	public void printUsage() {
		System.out.print("Usage : java -jar ElementClassify.jar ");
		System.out.print("[--element-table elementTable.xls] ");
		System.out.print("[--result result.xls] ");
		System.out.print("[--column-mapping columnMapping.xls] ");
		System.out.print("source01.xls source02.xls ...\n");
		System.exit(-255);
	}
	
	private boolean isValidExcelFilename(String filename) {
		if (filename.toLowerCase().endsWith(".xls")) { 
//				|| filename.toLowerCase().endsWith(".xlsx")) {
			return true;
		}
		if (filename.toLowerCase().endsWith(".xlsx")) {
			System.err.println(".xlsx format is not support yet");
		}
		return false;
	}
	
	public void parseArgument(String args[]) {
		if (args == null) {
			printUsage();
			return;
		}
		for (int i = 0; i < args.length; i++) {
			if (args[i].equalsIgnoreCase("--element-table") && i < args.length - 1) {
				++i;
				if (!isValidExcelFilename(args[i])) {
					System.err.println("Element table filename is not valid : " + args[i]);
					System.exit(-4);
				}
				elementTableFilename = args[i];
				System.out.println("Element table file name is " + elementTableFilename);
				continue;
			}
			if (args[i].equalsIgnoreCase("--column-mapping") && i < args.length - 1) {
				++i;
				if (!isValidExcelFilename(args[i])) {
					System.err.println("Column mapping filename is not valid : " + args[i]);
					System.exit(-4);
				}
				columnMappingFilename = args[i];
				System.out.println("Column mapping file name is " + columnMappingFilename);
				continue;
			}
			if (args[i].equalsIgnoreCase("--result") && i < args.length - 1) {
				++i;
				if (!isValidExcelFilename(args[i])) {
					System.err.println("Result filename is not valid : " + args[i]);
					System.out.println("Use result.xls as result filename");
				} else {
					resultFilename = args[i];
				}
				System.out.println("Result file name is " + resultFilename);
				continue;
			}
			if (!isValidExcelFilename(args[i])) {
				System.err.println("Ignore invalid source filename : " + args[i]);
			} else {
				sourceFilenames.add(args[i]);
			}
		}
		if (sourceFilenames.isEmpty()) {
			System.out.println("no source files");
			printUsage();
		}
	}
	
	ElementProperty[] getElementTable() {
		return elements;
	}
	ElementProperty[] elements;

	public void readElementTable(String filename) {
		// open elementTableFile and save element to ElementTable array
		Workbook elementBook = null;
		try {
			elementBook = Workbook.getWorkbook(new File(filename));
		} catch (Exception e) {
			e.printStackTrace();
		}

		if (elementBook == null) {
			System.err.println("open element table file failed!");
			System.exit(-5);
		}

		Sheet sheet = elementBook.getSheet(0); // get the fist sheet

		int rowNum = sheet.getRows();

		int columnNum = sheet.getColumns();

		ElementProperty[] element = new ElementProperty[rowNum - 1];

		int xMaxColumn = -1;
		int xMinColumn = -1;
		int yMaxColumn = -1;
		int yMinColumn = -1;

		for (int i = 1; i < columnNum; i++) {
			String cellCnt = sheet.getCell(i, 0).getContents();
			// System.out.println(cellCnt);
			if (cellCnt.compareToIgnoreCase("xMax") == 0) {
				xMaxColumn = i;
			} else if (cellCnt.compareToIgnoreCase("xMin") == 0) {
				xMinColumn = i;
			} else if (cellCnt.compareToIgnoreCase("yMax") == 0) {
				yMaxColumn = i;
			} else if (cellCnt.compareToIgnoreCase("yMin") == 0) {
				yMinColumn = i;
			}

			// System.out.println("xMax="+xMaxColumn+"xMin="+xMinColumn+"yMax="+yMaxColumn+"yMin="+yMinColumn);
		}

		for (int i = 1; i < rowNum; i++) {
			int index = i - 1;
			element[index] = new ElementProperty();
			element[index].elementName = sheet.getCell(0, i).getContents();
			element[index].xMax = Double.parseDouble(sheet.getCell(xMaxColumn,
					i).getContents());
			element[index].xMin = Double.parseDouble(sheet.getCell(xMinColumn,
					i).getContents());
			element[index].yMax = Double.parseDouble(sheet.getCell(yMaxColumn,
					i).getContents());
			element[index].yMin = Double.parseDouble(sheet.getCell(yMinColumn,
					i).getContents());
			// System.out.println("element["+index+"].elementName="+element[index].elementName+"element["+index+"].xMax="+element[index].xMax);
			// System.out.println("element["+index+"].yMin="+element[index].yMin+"element["+index+"].yMax="+element[index].yMax);
		}

		elementBook.close();
		
		elements = element;
	}
	
	public boolean writeResult() {
		// sample data write to new excel file
		ElementProperty[] element = elements;
		System.out.println("Writing result to " + getResultFilename());
		try {
			WritableWorkbook book = Workbook.createWorkbook(new File(getResultFilename()));
			WritableSheet wrtSheet = book.createSheet("the first page", 0);

			for (int i = 0; i < element.length; i++) {
				Label label = new Label(0, i + 1, element[i].elementName);
				wrtSheet.addCell(label);
			}

			for (int i = 0; i < sampleName.length; i++) {
				Label label = new Label(i + 1, 0, sampleName[i]);
				wrtSheet.addCell(label);
			}

			for (int i = 0; i < element.length; i++) {
				for (int j = 0; j < sampleName.length; j++) {
					Label label = new Label(j + 1, i + 1, value[i][j]);
					wrtSheet.addCell(label);
				}
			}
			book.write();
			book.close();

		} catch (Exception e) {
			System.out.println(e);
		}
		return true;
	}
	
	String xColumnName = "x";
	String yColumnName = "y";
	String valueColumnName = "value";
	String columnMappingFilename  = "columnMapping.xls";
	
	public boolean parseColumnMappingFile() {
		// open elementTableFile and save element to ElementTable array
		String filename = columnMappingFilename;
		Workbook mappingbook = null;
		try {
			mappingbook = Workbook.getWorkbook(new File(filename));
		} catch (Exception e) {
			e.printStackTrace();
		}

		if (mappingbook == null) {
			System.err.println("open columnMapping file failed!");
			return false;
		}

		Sheet sheet = mappingbook.getSheet(0); // get the fist sheet

		int rowNum = sheet.getRows();

		int columnNum = sheet.getColumns();
		
		for (int i = 0; i < rowNum; i++) {
			String appName = sheet.getCell(0, i).getContents();
			String toolName = sheet.getCell(1, i).getContents();
//			System.out.println(appName + "-" + toolName);
			if (appName.compareToIgnoreCase("x") == 0) {
				xColumnName = toolName;
			}
			else if (appName.compareToIgnoreCase("y") == 0) {
				yColumnName = toolName;
			}
			else if (appName.compareToIgnoreCase("value") == 0) {
				valueColumnName = toolName;
			}
		}

		mappingbook.close();

		return true;
	}
	
	public boolean parseSourceFiles() {
		parseColumnMappingFile();
		// open sample file, read all data and record all data
		ElementProperty [] element = elements;
		Vector<String> sources = getSourceFilenames();
		sampleName = new String[sources.size()];
		value = new String[element.length][sampleName.length];

		for (int i = 0; i < sources.size(); i++) {
			String filename = sources.get(i);
			System.out.println("source filename = " + filename);
			String tempS[] = filename.split("\\.");
			sampleName[i] = tempS[0];
			// System.out.println("tempS[0]="+tempS[0]+"  sampleName["+i+"]="+sampleName[i-1]);

			try {
				Workbook book = Workbook.getWorkbook(new File(filename));
				Sheet sheet = book.getSheet(0);
				int rowNum = sheet.getRows();
				int columnNum = sheet.getColumns();

				int xColumn = -1;
				int yColumn = -1;
				int valueColumn = -1;

				for (int j = 0; j < columnNum; j++) {
					String cellCnt = sheet.getCell(j, 0).getContents();
					// System.out.println(cellCnt);
					if (cellCnt.compareToIgnoreCase(xColumnName) == 0) {
						xColumn = j;
					} else if (cellCnt.compareToIgnoreCase(yColumnName) == 0) {
						yColumn = j;
					} else if (cellCnt.compareToIgnoreCase(valueColumnName) == 0) {
						valueColumn = j;
					}
					// System.out.println("x="+xColumn+"y="+yColumn+"value="+valueColumn);
				}
				if (xColumn == -1 || yColumn == -1 || valueColumn == -1) {
					System.err.println("column name is not match. " + xColumnName + " = " + xColumn + 
							", " + yColumnName + " = " + yColumn + ", " + 
							valueColumnName + " = " + valueColumn + ". @" + filename);
					continue;
				}

				for (int j = 1; j < rowNum; j++) {
					double x = Double.parseDouble(sheet.getCell(xColumn, j)
							.getContents());
					double y = Double.parseDouble(sheet.getCell(yColumn, j)
							.getContents());

					for (int k = 0; k < element.length; k++) {
						ElementProperty eTemp = element[k];
						if ((x >= eTemp.xMin) && (x <= eTemp.xMax)
								&& (y >= eTemp.yMin) && (y <= eTemp.yMax)) {
							value[k][i] = sheet.getCell(valueColumn, j)
									.getContents();
						}
					}
				}
				book.close();
			} catch (Exception e) {
				e.printStackTrace();
				System.out.println(e);
				System.exit(-3);
			}
		}
		return true;
	}
	
	String[] sampleName;
	String[][] value;
	
	public static void main(String args[]) {
		System.out.println(System.getProperty("user.dir"));
		EntryApp app = new EntryApp();
		app.parseArgument(args);

		// 1. read element table
		// open elementTableFile and save element to ElementTable array
		app.readElementTable(app.getElementTableFilename());
		ElementProperty[] element = app.getElementTable();

		// 2. classification
		app.parseSourceFiles();
//		// open sample file, read all data and record all data
//		Vector<String> sources = app.getSourceFilenames();
//		String[] sampleName = new String[sources.size()];
//		String[][] value = new String[element.length][sampleName.length];
//
//		for (int i = 0; i < sources.size(); i++) {
//			String filename = sources.get(i);
//			System.out.println("source filename = " + filename);
//			String tempS[] = filename.split("\\.");
//			sampleName[i] = tempS[0];
//			// System.out.println("tempS[0]="+tempS[0]+"  sampleName["+i+"]="+sampleName[i-1]);
//
//			try {
//				Workbook book = Workbook.getWorkbook(new File(filename));
//				Sheet sheet = book.getSheet(0);
//				int rowNum = sheet.getRows();
//				int columnNum = sheet.getColumns();
//
//				int xColumn = -1;
//				int yColumn = -1;
//				int valueColumn = -1;
//
//				for (int j = 0; j < columnNum; j++) {
//					String cellCnt = sheet.getCell(j, 0).getContents();
//					// System.out.println(cellCnt);
//					if (cellCnt.compareToIgnoreCase("x") == 0) {
//						xColumn = j;
//					} else if (cellCnt.compareToIgnoreCase("y") == 0) {
//						yColumn = j;
//					} else if (cellCnt.compareToIgnoreCase("value") == 0) {
//						valueColumn = j;
//					}
//					// System.out.println("x="+xColumn+"y="+yColumn+"value="+valueColumn);
//				}
//				if (xColumn == -1 || yColumn == -1 || valueColumn == -1) {
//					System.err.println("column name is not match. xColumn = " + xColumn + ", yColumn = " + yColumn + ", valueColumn = " + valueColumn + ". @" + filename);
//					continue;
//				}
//
//				for (int j = 1; j < rowNum; j++) {
//					double x = Double.parseDouble(sheet.getCell(xColumn, j)
//							.getContents());
//					double y = Double.parseDouble(sheet.getCell(yColumn, j)
//							.getContents());
//
//					for (int k = 0; k < element.length; k++) {
//						ElementProperty eTemp = element[k];
//						if ((x >= eTemp.xMin) && (x <= eTemp.xMax)
//								&& (y >= eTemp.yMin) && (y <= eTemp.yMax)) {
//							value[k][i] = sheet.getCell(valueColumn, j)
//									.getContents();
//						}
//					}
//				}
//				book.close();
//			} catch (Exception e) {
//				e.printStackTrace();
//				System.out.println(e);
//				System.exit(-3);
//			}
//		}

		// 3. write result
		app.writeResult();
//		// sample data write to new excel file
//		System.out.println("Writing result to " + app.getResultFilename());
//		try {
//			WritableWorkbook book = Workbook.createWorkbook(new File(
//					app.getResultFilename()));
//			WritableSheet wrtSheet = book.createSheet("the first page", 0);
//
//			for (int i = 0; i < element.length; i++) {
//				Label label = new Label(0, i + 1, element[i].elementName);
//				wrtSheet.addCell(label);
//			}
//
//			for (int i = 0; i < sampleName.length; i++) {
//				Label label = new Label(i + 1, 0, sampleName[i]);
//				wrtSheet.addCell(label);
//			}
//
//			for (int i = 0; i < element.length; i++) {
//				for (int j = 0; j < sampleName.length; j++) {
//					Label label = new Label(j + 1, i + 1, value[i][j]);
//					wrtSheet.addCell(label);
//				}
//			}
//			book.write();
//			book.close();
//
//		} catch (Exception e) {
//			System.out.println(e);
//		}
	}

	// @Override
	public void uncaughtException(Thread t, Throwable info) {
		info.printStackTrace();
	}

}
