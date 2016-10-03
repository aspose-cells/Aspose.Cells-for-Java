package com.aspose.cells.examples.data;

import com.aspose.cells.Cell;
import com.aspose.cells.Cells;
import com.aspose.cells.Style;
import com.aspose.cells.TextAlignmentType;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class TextAlignmentHorizontal {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(TextAlignmentHorizontal.class) + "data/";
		// Instantiating a Workbook object
		Workbook workbook = new Workbook();

		// Accessing the added worksheet in the Excel file
		int sheetIndex = workbook.getWorksheets().add();
		Worksheet worksheet = workbook.getWorksheets().get(sheetIndex);
		Cells cells = worksheet.getCells();

		// Adding the current system date to "A1" cell
		Cell cell = cells.get("A1");
		Style style = cell.getStyle();

		// Adding some value to the "A1" cell
		cell.setValue("Visit Aspose!");

		// Setting the horizontal alignment of the text in the "A1" cell
		style.setHorizontalAlignment(TextAlignmentType.CENTER);

		// Saved style
		cell.setStyle(style);

		// Saving the modified Excel file in default format
		workbook.save(dataDir + "TAHorizontal_out.xls");
	}
}
