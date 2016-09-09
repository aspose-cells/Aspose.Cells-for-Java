package com.aspose.cells.examples.files.handling;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;

public class SavingFiles {

	public static void main(String[] args) throws Exception {
		// ExStart:SavingFiles
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(SavingFiles.class) + "files/";

		// Creating an Workbook object with an Excel file path
		Workbook workbook = new Workbook(dataDir + "book.xls");

		// Save in default (Excel2003) format
		workbook.save(dataDir + "SavingFiles-default-out.xls");

		// Save in Excel2003 format
		workbook.save(dataDir + "SavingFiles03-out.xls", FileFormatType.EXCEL_97_TO_2003);

		// Save in Excel2007 xlsx format
		workbook.save(dataDir + "SavingFiles07-out.xlsx", FileFormatType.XLSX);

		// Save in SpreadsheetML format
		workbook.save(dataDir + "SavingFiles-out.xml", FileFormatType.EXCEL_2003_XML);

		// Print Message
		System.out.println("Worksheets are saved successfully.");
		// ExEnd:SavingFiles
	}
}
