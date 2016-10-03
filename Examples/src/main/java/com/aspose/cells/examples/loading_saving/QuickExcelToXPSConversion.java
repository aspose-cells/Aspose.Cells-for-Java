package com.aspose.cells.examples.loading_saving;

import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class QuickExcelToXPSConversion {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ExportWholeWorkbookToXPS.class) + "loading_saving/";
		// Open an Excel file
		Workbook workbook = new Workbook(dataDir + "Book1.xlsx");

		// Save in XPS format
		workbook.save("QEToXPSConversion_out.xps", SaveFormat.XPS);
	}
}
