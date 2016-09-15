package com.aspose.cells.examples.worksheets.security.protecting;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;

public class ProtectingSpecificRowInWorksheet {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ProtectingSpecificRowInWorksheet.class) + "worksheets/";

		// Create a new workbook.
		Workbook wb = new Workbook();

		// obtain the first sheet.
		Worksheet sheet = wb.getWorksheets().get(0);

		// Define the style object.
		Style style;

		// Define the styleflag object.
		StyleFlag flag;

		// Loop through all the columns in the worksheet and unlock them.
		for (int i = 0; i <= 255; i++) {
			style = sheet.getCells().getColumns().get(i).getStyle();
			style.setLocked(false);
			flag = new StyleFlag();
			flag.setLocked(true);
			sheet.getCells().getColumns().get(i).applyStyle(style, flag);
		}

		// Get the first row style.
		style = sheet.getCells().getRows().get(0).getStyle();

		// Lock it.
		style.setLocked(true);

		// Instantiate the flag.
		flag = new StyleFlag();

		// Set the lock setting.
		flag.setLocked(true);

		// Apply the style to the first row.
		sheet.getCells().getRows().get(0).applyStyle(style, flag);

		// Save the excel file.
		wb.save(dataDir + "PSRowInWorksheet-out.xls", FileFormatType.EXCEL_97_TO_2003);

		// Print Message
		System.out.println("Row protected successfully.");

	}
}
