package com.aspose.cells.examples.articles;

import com.aspose.cells.Cells;
import com.aspose.cells.Color;
import com.aspose.cells.Range;
import com.aspose.cells.Style;
import com.aspose.cells.StyleFlag;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class CreatingStyle {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(CreatingStyle.class) + "articles/";
		// Create a workbook.
		Workbook workbook = new Workbook();

		// Create a new style object.
		int styleIdx = workbook.getStyles().add();
		Style style = workbook.getStyles().get(styleIdx);

		// Set the number format.
		style.setNumber(14);

		// Set the font color to red color.
		style.getFont().setColor(Color.getRed());

		// Name the style.
		style.setName("Date1");

		// Get the first worksheet cells.
		Cells cells = workbook.getWorksheets().get(0).getCells();

		// Specify the style (described above) to A1 cell.
		cells.get("A1").setStyle(style);

		// Create a range (B1:D1).
		Range range = cells.createRange("B1", "D1");

		// Initialize styleflag object.
		StyleFlag flag = new StyleFlag();

		// Set all formatting attributes on.
		flag.setAll(true);

		// Apply the style (described above)to the range.
		range.applyStyle(style, flag);

		// Modify the style (described above) and change the font color from red to black.
		style.getFont().setColor(Color.getBlack());

		// Done! Since the named style (described above) has been set to a cell and range,the change would be Reflected(new
		// modification is implemented) to cell(A1) and //range (B1:D1).
		style.update();

		// Save the excel file.
		workbook.save(dataDir + "CreatingStyle_out.xls");

	}
}
