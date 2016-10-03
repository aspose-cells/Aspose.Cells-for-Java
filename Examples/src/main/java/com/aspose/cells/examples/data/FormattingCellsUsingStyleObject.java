package com.aspose.cells.examples.data;

import com.aspose.cells.BorderType;
import com.aspose.cells.Cell;
import com.aspose.cells.CellBorderType;
import com.aspose.cells.Cells;
import com.aspose.cells.Color;
import com.aspose.cells.Font;
import com.aspose.cells.Style;
import com.aspose.cells.TextAlignmentType;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class FormattingCellsUsingStyleObject {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(FormattingCellsUsingStyleObject.class) + "data/";
		// Instantiating a Workbook object
		Workbook workbook = new Workbook(dataDir + "book1.xls");

		// Accessing the first worksheet in the Excel file
		Worksheet worksheet = workbook.getWorksheets().get(0);
		Cells cells = worksheet.getCells();

		// Accessing the "A1" cell from the worksheet
		Cell cell = cells.get("A1");

		// Adding some value to the "A1" cell
		cell.setValue("Hello Aspose!");

		// Adding a new Style to the styles collection of the Excel object
		Style style = workbook.createStyle();

		// Setting the vertical alignment of the text in the "A1" cell
		style.setVerticalAlignment(TextAlignmentType.CENTER);

		// Setting the horizontal alignment of the text in the "A1" cell
		style.setHorizontalAlignment(TextAlignmentType.CENTER);

		// Setting the font color of the text in the "A1" cell
		Font font = style.getFont();
		font.setColor(Color.getGreen());

		// Setting the cell to shrink according to the text contained in it
		style.setShrinkToFit(true);

		// Setting the bottom border
		style.setBorder(BorderType.BOTTOM_BORDER, CellBorderType.MEDIUM, Color.getRed());

		// Saved style
		cell.setStyle(style);

		// Saving the modified Excel file in default format
		workbook.save(dataDir + "FCUsingStyleObject_out.xls");
	}
}
