package com.aspose.cells.examples.data;

import com.aspose.cells.Cell;
import com.aspose.cells.Color;
import com.aspose.cells.Font;
import com.aspose.cells.Style;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class ColorsAndPalette {
	public static void main(String[] args) throws Exception {
		// Path to source file
		String dataDir = Utils.getSharedDataDir(ColorsAndPalette.class) + "data/";
		// Instantiating a Workbook object
		Workbook workbook = new Workbook();

		// Adding custom color to the palette at 55th index
		Color color = Color.fromArgb(212, 213, 0);
		workbook.changePalette(color, 55);

		// Obtaining the reference of the newly added worksheet by passing its sheet index
		int sheetIndex = workbook.getWorksheets().add();
		Worksheet worksheet = workbook.getWorksheets().get(sheetIndex);

		// Accessing the "A1" cell from the worksheet
		Cell cell = worksheet.getCells().get("A1");

		// Adding some value to the "A1" cell
		cell.setValue("Hello Aspose!");

		// Setting the custom color to the font
		Style style = cell.getStyle();
		Font font = style.getFont();
		font.setColor(color);

		cell.setStyle(style);

		// Saving the Excel file
		workbook.save(dataDir + "ColorsAndPalette_out.xls");
	}
}
