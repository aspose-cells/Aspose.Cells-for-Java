package com.aspose.cells.examples.loading_saving;

import com.aspose.cells.FileFormatType;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class SavingFiletoSomeLocation {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(SavingFiletoSomeLocation.class) + "loading_saving/";

		String filePath = dataDir + "Book1.xls";

		// Creating an Workbook object with an Excel file path
		Workbook workbook = new Workbook(filePath);

		// Save in Excel 97 – 2003 format
		workbook.save(dataDir + "SFTSomeLocation_out.xls");
		// OR
		// workbook.save(dataDir + ".output..xls", new
		// XlsSaveOptions(SaveFormat.Excel97To2003));

		// Save in Excel2007 xlsx format
		workbook.save(dataDir + "SFTSomeLocation_out.xlsx", FileFormatType.XLSX);

		// Save in Excel2007 xlsb format
		workbook.save(dataDir + "SFTSomeLocation_out.xlsb", FileFormatType.XLSB);

		// Save in ODS format
		workbook.save(dataDir + "SFTSomeLocation_out.ods", FileFormatType.ODS);

		// Save in Pdf format
		workbook.save(dataDir + "SFTSomeLocation_out.pdf", FileFormatType.PDF);

		// Save in Html format
		workbook.save(dataDir + "SFTSomeLocation_out.html", FileFormatType.HTML);

		// Save in SpreadsheetML format
		workbook.save(dataDir + "SFTSomeLocation_out.xml", FileFormatType.EXCEL_2003_XML);

		// Print Message
		System.out.println("Worksheets are saved successfully.");

	}
}
