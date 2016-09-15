package com.aspose.cells.examples.files.handling;

import com.aspose.cells.FileFormatType;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class SaveFileInExcel972003format {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(SaveFileInExcel972003format.class) + "files/";

		// Creating an Workbook object with an Excel file path
		Workbook workbook = new Workbook();

		// Save in Excel2003 format
		workbook.save(dataDir + "SFIExcel972003format-out.xls", FileFormatType.EXCEL_97_TO_2003);

		// Print Message
		System.out.println("Worksheets are saved successfully.");

	}
}
