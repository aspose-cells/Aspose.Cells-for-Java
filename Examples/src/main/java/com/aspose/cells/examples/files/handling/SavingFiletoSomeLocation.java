package com.aspose.cells.examples.files.handling;

import com.aspose.cells.FileFormatType;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class SavingFiletoSomeLocation {

	public static void main(String[] args) throws Exception {
		// ExStart:1
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(SavingFiletoSomeLocation.class) + "files/";

		String filePath = dataDir + "Book1.xls";

		// Creating an Workbook object with an Excel file path
		Workbook workbook = new Workbook(filePath);

		// Save in Excel 97 – 2003 format
		workbook.save(dataDir + ".output.xls");
		// OR
		// workbook.save(dataDir + ".output..xls", new
		// XlsSaveOptions(SaveFormat.Excel97To2003));

		// Save in Excel2007 xlsx format
		workbook.save(dataDir + "SFTSomeLocation-out.xlsx", FileFormatType.XLSX);

		// Save in Excel2007 xlsb format
		workbook.save(dataDir + "SFTSomeLocation-out.xlsb", FileFormatType.XLSB);

		// Save in ODS format
		workbook.save(dataDir + "SFTSomeLocation-out.ods", FileFormatType.ODS);

		// Save in Pdf format
		workbook.save(dataDir + "SFTSomeLocation-out.pdf", FileFormatType.PDF);

		// Save in Html format
		workbook.save(dataDir + "SFTSomeLocation-out.html", FileFormatType.HTML);

		// Save in SpreadsheetML format
		workbook.save(dataDir + "SFTSomeLocation-out.xml", FileFormatType.EXCEL_2003_XML);

		// Print Message
		System.out.println("Worksheets are saved successfully.");
		// ExEnd:1
	}
}
