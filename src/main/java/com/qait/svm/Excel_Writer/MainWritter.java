package com.qait.svm.Excel_Writer;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class MainWritter {

	public static final String DESTINATION_FILE_NAME = "java_demo.xlsx";

	public static void main(String[] args) throws IOException {

		FileOutputStream out = new FileOutputStream(new File(DESTINATION_FILE_NAME));
		Workbook workbook = ExcelWritter.getWorkbook(out, DESTINATION_FILE_NAME);
		Sheet sheet = workbook.createSheet(DESTINATION_FILE_NAME);

		ExcelWritter.write(sheet);

		workbook.write(out);
		out.close();
		workbook.close();
		System.out.println("Hey! Your file " + DESTINATION_FILE_NAME + " has been written successfully on disk.");
	}
}
