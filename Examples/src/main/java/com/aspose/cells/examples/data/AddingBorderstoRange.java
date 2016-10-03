package com.aspose.cells.examples.data;

import com.aspose.cells.Cell;
import com.aspose.cells.CellBorderType;
import com.aspose.cells.Color;
import com.aspose.cells.Range;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class AddingBorderstoRange {
	public static void main(String[] args) throws Exception {
		// Path to source file
		String dataDir = Utils.getSharedDataDir(AddingBordersToCells.class) + "data/";
		// Instantiating a Workbook object
		Workbook workbook = new Workbook();

		// Adding a new worksheet to the Workbook object Obtaining the reference of the newly added worksheet
		int sheetIndex = workbook.getWorksheets().add();
		Worksheet worksheet = workbook.getWorksheets().get(sheetIndex);

		// Accessing the "A1" cell from the worksheet
		Cell cell = worksheet.getCells().get("A1");

		// Adding some value to the "A1" cell
		cell.setValue("Hello World From Aspose");

		// Creating a range of cells starting from "A1" cell to 3rd column in a
		// row
		Range range = worksheet.getCells().createRange(0, 0, 1, 2);
		range.setName("MyRange");

		// Adding a thick outline border with the blue line
		range.setOutlineBorders(CellBorderType.THICK, Color.getBlue());

		// Saving the Excel file
		workbook.save(dataDir + "ABToRange_out.xls");
	}
}
