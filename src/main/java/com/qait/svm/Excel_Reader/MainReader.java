package com.qait.svm.Excel_Reader;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
/**
 * 
 * @author shivambharadwaj
 *
 */
public class MainReader {

	private static final String FILE_PATH = "Book1.xls";

	public static void main(String[] args) throws IOException {
		
		FileInputStream is = new FileInputStream(new File(FILE_PATH));
		Workbook woorkbook = MyWorkBook.getWorkbook(is, FILE_PATH);
		// Getting sheet by index
		Sheet sheet = woorkbook.getSheetAt(0);
		if (sheet.isSelected()) {
			Iterator<Row> rows = sheet.rowIterator();
			MyWorkBook.printSheetToConsole(rows);
		} else {
			System.out.println("select the sheet first...");
		}
		woorkbook.close();
		is.close();
	}
	
}
