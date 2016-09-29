package com.aspose.cells.examples.articles;

import com.aspose.cells.PasteOptions;
import com.aspose.cells.PasteType;
import com.aspose.cells.Range;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class CopyRowHeights {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(CopyRowHeights.class) + "articles/";
		// Create workbook object
		Workbook workbook = new Workbook();

		// Source worksheet
		Worksheet srcSheet = workbook.getWorksheets().get(0);

		// Add destination worksheet
		Worksheet dstSheet = workbook.getWorksheets().add("Destination Sheet");

		// Set the row height of the 4th row
		// This row height will be copied to destination range
		srcSheet.getCells().setRowHeight(3, 50);

		// Create source range to be copied
		Range srcRange = srcSheet.getCells().createRange("A1:D10");

		// Create destination range in destination worksheet
		Range dstRange = dstSheet.getCells().createRange("A1:D10");

		// PasteOptions, we want to copy row heights of source range to destination range
		PasteOptions opts = new PasteOptions();
		opts.setPasteType(PasteType.ROW_HEIGHTS);

		// Copy source range to destination range with paste options
		dstRange.copy(srcRange, opts);

		// Write informative message in cell D4 of destination worksheet
		dstSheet.getCells().get("D4").putValue("Row heights of source range copied to destination range");

		// Save the workbook in xlsx format
		workbook.save(dataDir + "CopyRowHeights_out.xlsx", SaveFormat.XLSX);

	}
}
