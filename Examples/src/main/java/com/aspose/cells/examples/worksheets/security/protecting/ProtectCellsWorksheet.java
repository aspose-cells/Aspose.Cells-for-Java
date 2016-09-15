package com.aspose.cells.examples.worksheets.security.protecting;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;

public class ProtectCellsWorksheet {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ProtectCellsWorksheet.class) + "worksheets/";

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

		// Lock the three cells...i.e. A1, B1, C1.
		style = sheet.getCells().get("A1").getStyle();
		style.setLocked(true);
		sheet.getCells().get("A1").setStyle(style);

		style = sheet.getCells().get("B1").getStyle();
		style.setLocked(true);
		sheet.getCells().get("B1").setStyle(style);

		style = sheet.getCells().get("C1").getStyle();
		style.setLocked(true);
		sheet.getCells().get("C1").setStyle(style);

		sheet.protect(ProtectionType.ALL);

		// Save the excel file.
		wb.save(dataDir + "PCellsWorksheet-out.xls", FileFormatType.EXCEL_97_TO_2003);

		// Print Message
		System.out.println("Cell  protected successfully.");

	}
}
