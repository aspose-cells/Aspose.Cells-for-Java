package com.aspose.cells.examples.worksheets.security;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;

public class LockCell {

	public static void main(String[] args) throws Exception {
		// ExStart:1
		// The path to the documents directory.
		String dataDir = Utils.getDataDir(LockCell.class);

		// Instantiating a Workbook object by excel file path
		Workbook excel = new Workbook(dataDir + "Book1.xlsx");

		WorksheetCollection worksheets = excel.getWorksheets();
		Worksheet worksheet = worksheets.get(0);

		worksheet.getCells().get("A1").getStyle().setLocked(true);

		// Saving the modified Excel file Excel XP format
		excel.save(dataDir + "output.xls", FileFormatType.EXCEL_97_TO_2003);

		// Print Message
		System.out.println("Cell Locked successfully.");
		// ExEnd:1
	}
}
