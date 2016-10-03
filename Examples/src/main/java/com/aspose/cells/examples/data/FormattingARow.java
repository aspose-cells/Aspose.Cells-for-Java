package com.aspose.cells.examples.data;

import com.aspose.cells.BorderType;
import com.aspose.cells.CellBorderType;
import com.aspose.cells.Cells;
import com.aspose.cells.Color;
import com.aspose.cells.Font;
import com.aspose.cells.Row;
import com.aspose.cells.Style;
import com.aspose.cells.StyleFlag;
import com.aspose.cells.TextAlignmentType;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class FormattingARow {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(FormattingARow.class) + "data/";
		// Instantiating a Workbook object
		Workbook workbook = new Workbook();

		// Accessing the added worksheet in the Excel file
		Worksheet worksheet = workbook.getWorksheets().get(0);
		Cells cells = worksheet.getCells();

		// Adding a new Style to the styles collection of the Excel object Accessing the newly added Style to the Excel object
		Style style = workbook.createStyle();

		// Setting the vertical alignment of the text in the cell
		style.setVerticalAlignment(TextAlignmentType.CENTER);

		// Setting the horizontal alignment of the text in the cell
		style.setHorizontalAlignment(TextAlignmentType.CENTER);

		// Setting the font color of the text in the cell
		Font font = style.getFont();
		font.setColor(Color.getGreen());

		// Shrinking the text to fit in the cell
		style.setShrinkToFit(true);

		// Setting the bottom border of the cell
		style.setBorder(BorderType.BOTTOM_BORDER, CellBorderType.MEDIUM, Color.getRed());

		// Creating StyleFlag
		StyleFlag styleFlag = new StyleFlag();
		styleFlag.setHorizontalAlignment(true);
		styleFlag.setVerticalAlignment(true);
		styleFlag.setShrinkToFit(true);
		styleFlag.setBottomBorder(true);
		styleFlag.setFontColor(true);

		// Accessing a row from the Rows collection
		Row row = cells.getRows().get(0);

		// Assigning the Style object to the Style property of the row
		row.applyStyle(style, styleFlag);

		// Saving the Excel file
		workbook.save(dataDir + "FormattingARow_out.xls");
	}
}
