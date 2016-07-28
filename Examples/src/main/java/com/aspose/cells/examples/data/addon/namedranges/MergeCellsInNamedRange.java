package com.aspose.cells.examples.data.addon.namedranges;

import com.aspose.cells.Range;
import com.aspose.cells.Style;
import com.aspose.cells.StyleFlag;
import com.aspose.cells.TextAlignmentType;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class MergeCellsInNamedRange {

	public static void main(String[] args) throws Exception {
		// ExStart:MergeCellsInNamedRange
		// The path to the documents directory.
		String dataDir = Utils.getDataDir(MergeCellsInNamedRange.class);

		// Instantiate a new Workbook.
		Workbook wb1 = new Workbook();

		// Get the first worksheet in the workbook.
		Worksheet worksheet1 = wb1.getWorksheets().get(0);

		// Create a range.
		Range mrange = worksheet1.getCells().createRange("A18", "J18");

		// Name the range.
		mrange.setName("Details");

		// Merge the cells of the range.
		mrange.merge();

		// Get the range.
		Range range1 = wb1.getWorksheets().getRangeByName("Details");

		// Define a style object.
		Style style = wb1.createStyle();

		// Set the alignment.
		style.setHorizontalAlignment(TextAlignmentType.CENTER);

		// Create a StyleFlag object.
		StyleFlag flag = new StyleFlag();
		// Make the relative style attribute ON.
		flag.setHorizontalAlignment(true);

		// Apply the style to the range.
		range1.applyStyle(style, flag);

		// Input data into range.
		range1.get(0, 0).setValue("Aspose");

		// Save the excel file.
		wb1.save(dataDir + "mergingrange.out.xls");

		// Print message
		System.out.println("Process completed successfully");
		// ExEnd:MergeCellsInNamedRange
	}
}
