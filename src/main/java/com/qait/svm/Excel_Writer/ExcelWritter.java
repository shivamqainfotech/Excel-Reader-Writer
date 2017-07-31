package com.qait.svm.Excel_Writer;

import java.io.FileOutputStream;
import java.util.Map;
import java.util.Scanner;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
/**
 * 
 * @author shivambharadwaj
 *
 */
public class ExcelWritter {

	public static Workbook getWorkbook(FileOutputStream out, String excelFilePath) {
		Workbook workbook = null;

		if (excelFilePath.endsWith("xlsx")) {
			workbook = new XSSFWorkbook();

		} else if (excelFilePath.endsWith("xls")) {
			workbook = new HSSFWorkbook();

		} else {
			throw new IllegalArgumentException("The specified file is not Excel file");
		}
		return workbook;
	}

	public static Map<String, Object[]> takeUserInput() {
		Map<String, Object[]> data = new TreeMap<String, Object[]>();
		data.put("1", new Object[] { "ID", "NAME", "LASTNAME" });
		Scanner scan = new Scanner(System.in);
		System.err.println("Enter the details as per thr format given-------");
		System.err.println("ID\nName\nLastName");
		String id = null;
		String name = null;
		String lname = null;
		int k = 1;
		while (true) {
			id = scan.nextLine();
			name = scan.nextLine();
			lname = scan.nextLine();
			k = k + 1;
			data.put(String.valueOf(k), new Object[] { id, name, lname });
			System.err.println("Alert! Do you want to continue....(y/n)");

			String i = scan.nextLine();
			if (i.equalsIgnoreCase("n")) {
				break;
			}
		}
		return data;
	}

	public static void write(Sheet sheet) {
		Map<String, Object[]> data = takeUserInput();

		Set<String> keyset = data.keySet();// Getting{1,2,3,4,5,6}

		int rownum = 0;
		for (String key : keyset) {
			Row row = sheet.createRow(rownum++);

			Object[] objArr = data.get(key);

			int cellnum = 0;
			for (Object obj : objArr) {
				Cell cell = row.createCell(cellnum++);

				if (obj instanceof String)
					cell.setCellValue((String) obj);
				else if (obj instanceof Integer)
					cell.setCellValue((Integer) obj);
				else if (obj instanceof Boolean)
					cell.setCellValue((Boolean) obj);
				else if (obj instanceof Character)
					cell.setCellValue((Character) obj);
				else if (obj instanceof Double)
					cell.setCellValue((Double) obj);
			}
		}

	}

}
