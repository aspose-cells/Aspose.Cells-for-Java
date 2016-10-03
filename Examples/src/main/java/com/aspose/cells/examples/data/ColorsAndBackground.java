package com.aspose.cells.examples.data;

import com.aspose.cells.BackgroundType;
import com.aspose.cells.Cell;
import com.aspose.cells.Cells;
import com.aspose.cells.Color;
import com.aspose.cells.Style;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class ColorsAndBackground {
	public static void main(String[] args) throws Exception {
		String dataDir = Utils.getSharedDataDir(ColorsAndBackground.class) + "data/";
		// Instantiating a Workbook object
		Workbook workbook = new Workbook();

		// Accessing the added worksheet in the Excel file
		int sheetIndex = workbook.getWorksheets().add();
		Worksheet worksheet = workbook.getWorksheets().get(sheetIndex);
		Cells cells = worksheet.getCells();

		// Accessing the "A1" cell from the worksheet
		Cell cell = cells.get("A1");
		Style style = cell.getStyle();

		// Setting the foreground color to yellow
		style.setBackgroundColor(Color.getYellow());

		// Setting the background pattern to vertical stripe
		style.setPattern(BackgroundType.VERTICAL_STRIPE);

		// Saving the modified style to the "A1" cell.
		cell.setStyle(style);

		// Accessing the "A2" cell from the worksheet
		cell = cells.get("A2");
		style = cell.getStyle();

		// Setting the foreground color to blue
		style.setBackgroundColor(Color.getBlue());

		// Setting the background color to yellow
		style.setForegroundColor(Color.getYellow());

		// Setting the background pattern to vertical stripe
		style.setPattern(BackgroundType.VERTICAL_STRIPE);

		// Saving the modified style to the "A2" cell.
		cell.setStyle(style);

		// Saving the Excel file
		workbook.save(dataDir + "ColorsAndBackground_out.xls");
	}
}
