package ck;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Vector;

//import org.apache.poi.hssf.usermodel.HSSFRow;
//import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
//import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelFileReader implements ITableFileReader {
	
	ExcelFileReader reader;
	String filename;
	Vector<Vector<String>> data;
	
	public ExcelFileReader(String filename) {
		this.filename = filename;
		loadData();
	}

	public void loadData() {
		try {
			FileInputStream fis = new FileInputStream(filename);
			Workbook wb = null;
			ExcelFileType t = getFileType(filename);
			switch(t) {
			case XLS:
				wb = new HSSFWorkbook(fis);
				break;
			case XLSX:
				wb = new XSSFWorkbook(fis);
				break;
			default:
				System.err.println(filename + " is not excel file");
				fis.close();
				return;
			}
			Sheet sheet = wb.getSheetAt(0);		// get first sheet
			data = new Vector<Vector<String>>();
			System.out.println("total rows = " + sheet.getLastRowNum());
			for (int i = 0; i <= sheet.getLastRowNum(); i++) {		// fuck
				Vector<String> rowData = new Vector<String>();
				Row r = sheet.getRow(i);
				if (r == null) {
					System.err.println("" + i + "th row is null");
					continue;
				} 
				System.out.println("" + i + "th row cols = " + r.getLastCellNum());
				for (int j = 0; j < r.getLastCellNum(); j++) {
					Cell c = r.getCell(j);
					if (c == null) {
						System.err.println("" + i + "th row " + j + "th col is null");
						continue;
					}
					System.out.println("reading " + i + "th row " + j + "th col");
					c.setCellType(Cell.CELL_TYPE_STRING);
					String content = c.getStringCellValue();
					rowData.add(content);
				}
				data.add(rowData);
			}
			fis.close();
		} catch (FileNotFoundException fnfe) {
			fnfe.printStackTrace();
		} catch (IOException ioe) {
			ioe.printStackTrace();
		}
	}
	
	public Vector<Vector<String>> getData() {
		return data;
	}
	
	enum ExcelFileType {
		XLS,			// binary format, used for excel 97-2003
		XLSX,			// xml format, used since excel 2007
		INVALID			// not an excel file
	}

	private ExcelFileType getFileType(String filename) {
		if (filename.toLowerCase().endsWith(".xls")) {
			return ExcelFileType.XLS;
		}
		else if (filename.toLowerCase().endsWith(".xlsx")) {
			return ExcelFileType.XLSX;
		}
		return ExcelFileType.INVALID;
	}
	
	class BinaryExcelFileReader extends ExcelFileReader {
		public BinaryExcelFileReader(String filename) {
			super(filename);
		}
		
		public void loadData() {
			try {
				FileInputStream fis = new FileInputStream(filename);
				Workbook wb = new HSSFWorkbook(fis);
				Sheet sheet = wb.getSheetAt(0);		// get first sheet
				data = new Vector<Vector<String>>();
				for (int i = 0; i < sheet.getLastRowNum(); i++) {
					Vector<String> rowData = new Vector<String>();
					Row r = sheet.getRow(i);
					if (r == null) {
						System.err.println("" + i + "th row is null");
						continue;
					} 
					for (int j = 0; j < r.getLastCellNum(); j++) {
						Cell c = r.getCell(j);
						if (c == null) {
							System.err.println("" + i + "th row " + j + "th col is null");
							continue;
						}
						System.out.println("reading " + i + "th row " + j + "th col");
						c.setCellType(Cell.CELL_TYPE_STRING);
						String content = c.getStringCellValue();
						rowData.add(content);
					}
					data.add(rowData);
				}
			} catch (FileNotFoundException fnfe) {
				fnfe.printStackTrace();
			} catch (IOException ioe) {
				ioe.printStackTrace();
			}
		}
	}
	
	class XMLExcelFileReader extends ExcelFileReader {

		public XMLExcelFileReader(String filename) {
			super(filename);
		}
		
		public void loadData() {
			try {
				FileInputStream fis = new FileInputStream(filename);
				Workbook wb = new XSSFWorkbook(fis);
				Sheet sheet = wb.getSheetAt(0);
				data = new Vector<Vector<String>>();
				for (int i = 0; i <= sheet.getLastRowNum(); i++) {
					Vector<String> rowData = new Vector<String>();
					Row r = sheet.getRow(i);
					for (int j = 0; j <= r.getLastCellNum(); j++) {
						String content = r.getCell(j).getStringCellValue();
						rowData.add(content);
					}
					data.add(rowData);
				}
			} catch (FileNotFoundException fnfe) {
				fnfe.printStackTrace();
			} catch (IOException ioe) {
				ioe.printStackTrace();
			}
		}
		
	}

}
