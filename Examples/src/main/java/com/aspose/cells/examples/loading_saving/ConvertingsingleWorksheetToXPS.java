package com.aspose.cells.examples.loading_saving;

import com.aspose.cells.ImageOrPrintOptions;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.SheetRender;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class ConvertingsingleWorksheetToXPS {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ConvertingsingleWorksheetToXPS.class) + "loading_saving/";
		
		//Open an Excel file
		Workbook workbook = new Workbook(dataDir + "Book1.xlsx");

		//Get the first worksheet
		Worksheet sheet = workbook.getWorksheets().get(0);

		//Apply different Image and Print options
		ImageOrPrintOptions options = new ImageOrPrintOptions();
		//Set the format
		options.setSaveFormat(SaveFormat.XPS);

		//Render the sheet with respect to specified printing options
		SheetRender render = new SheetRender(sheet, options);
		render.toImage(0, dataDir + "CSingleWorksheetToXPS_out.xps");
	}
}
