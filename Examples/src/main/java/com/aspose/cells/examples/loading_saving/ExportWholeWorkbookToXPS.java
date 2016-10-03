package com.aspose.cells.examples.loading_saving;

import com.aspose.cells.ImageOrPrintOptions;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;
import com.aspose.cells.WorkbookRender;
import com.aspose.cells.examples.Utils;

public class ExportWholeWorkbookToXPS {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ExportWholeWorkbookToXPS.class) + "loading_saving/";
		// Open an Excel file
		Workbook workbook = new Workbook(dataDir + "Book1.xlsx");

		// Apply different Image and Print options
		ImageOrPrintOptions options = new ImageOrPrintOptions();
		// Set the format
		options.setSaveFormat(SaveFormat.XPS);

		// Render the workbook with respect to specified printing options
		WorkbookRender render = new WorkbookRender(workbook, options);
		render.toImage(dataDir + "ExportWholeWorkbookToXPS_out.xps");
	}
}
