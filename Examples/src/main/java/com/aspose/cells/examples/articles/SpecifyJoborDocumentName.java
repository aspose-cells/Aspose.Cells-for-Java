package com.aspose.cells.examples.articles;

import com.aspose.cells.ImageOrPrintOptions;
import com.aspose.cells.SheetRender;
import com.aspose.cells.Workbook;
import com.aspose.cells.WorkbookRender;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class SpecifyJoborDocumentName {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(SpecifyJoborDocumentName.class) + "articles/";
		// Create workbook object from source Excel file
		Workbook workbook = new Workbook(dataDir + "source.xlsx");

		// Specify Printer and Job Name
		String printerName = "doPDF v7";
		String jobName = "Job Name while Printing with Aspose.Cells";

		// Print workbook using WorkbookRender
		WorkbookRender wr = new WorkbookRender(workbook, new ImageOrPrintOptions());
		wr.toPrinter(printerName, jobName);

		// Access first worksheet
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Print worksheet using SheetRender
		SheetRender sr = new SheetRender(worksheet, new ImageOrPrintOptions());
		sr.toPrinter(printerName, jobName);


	}
}
