package com.aspose.cells.examples.data;

import com.aspose.cells.Cell;
import com.aspose.cells.Cells;
import com.aspose.cells.Style;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class Indentation {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(Indentation.class) + "data/";
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

		// Setting the vertical alignment of the text in a cell
		Style style1 = cell.getStyle();
		style1.setIndentLevel(2);
		cell.setStyle(style1);

		// Saved style
		cell.setStyle(style1);

		// Saving the modified Excel file in default format
		workbook.save(dataDir + "Indentation_out.xls");
	}
}
