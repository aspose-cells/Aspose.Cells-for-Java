package com.aspose.cells.examples.worksheets;

import com.aspose.cells.*;

import com.aspose.cells.examples.Utils;

public class ProtectingSpecificCellsinaWorksheet {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ProtectingSpecificCellsinaWorksheet.class) + "worksheets/";

		// Create a new workbook.
		Workbook wb = new Workbook();

		// Create a worksheet object and obtain the first sheet.
		Worksheet sheet = wb.getWorksheets().get(0);

		// Define the style object.
		Style style;

		// Define the styleflag object.
		StyleFlag flag;
		flag = new StyleFlag();
		flag.setLocked(true);

		// Loop through all the columns in the worksheet and unlock them.
		for (int i = 0; i <= 255; i++) {
			style = sheet.getCells().getColumns().get(i).getStyle();
			style.setLocked(false);
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

		// Save the excel file.
		wb.save(dataDir + "PSpecificCellsinaWorksheet_out.xls", FileFormatType.EXCEL_97_TO_2003);

		// Print Message
		System.out.println("Cells protected successfully.");

	}
}
