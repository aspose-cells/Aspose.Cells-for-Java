package com.aspose.cells.examples.loading_saving;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;

public class ConvertingToXPS {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ConvertingToXPS.class) + "loading_saving/";

		Workbook workbook = new Workbook(dataDir + "Book1.xls");

		// Get the first worksheet.
		Worksheet sheet = workbook.getWorksheets().get(0);

		// Apply different Image and Print options
		com.aspose.cells.ImageOrPrintOptions options = new ImageOrPrintOptions();

		// Set the Format
		options.setSaveFormat(SaveFormat.XPS);

		// Render the sheet with respect to specified printing options
		com.aspose.cells.SheetRender sr = new SheetRender(sheet, options);
		sr.toImage(0, dataDir + "ConvertingToXPS_out.xps");

		// Save the complete Workbook in XPS format
		workbook.save(dataDir + "ConvertingToXPS_out.xps", SaveFormat.XPS);

		// Print message
		System.out.println("Excel to XPS conversion performed successfully.");

	}
}
