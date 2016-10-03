package com.aspose.cells.examples.data;

import com.aspose.cells.Color;
import com.aspose.cells.GradientStyleType;
import com.aspose.cells.Style;
import com.aspose.cells.TextAlignmentType;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class ApplyGradientFillEffects {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ApplyGradientFillEffects.class) + "data/";
		// Instantiate a new Workbook
		Workbook workbook = new Workbook();
		// Get the first worksheet (default) in the workbook
		Worksheet worksheet = workbook.getWorksheets().get(0);
		// Input a value into B3 cell
		worksheet.getCells().get(2, 1).putValue("test");

		// Get the Style of the cell
		Style style = worksheet.getCells().get("B3").getStyle();
		// Set Gradient pattern on
		style.setGradient(true);
		// Specify two color gradient fill effects
		style.setTwoColorGradient(Color.fromArgb(255, 255, 255), Color.fromArgb(79, 129, 189),
				GradientStyleType.HORIZONTAL, 1);
		// Set the color of the text in the cell
		style.getFont().setColor(Color.getRed());
		// Specify horizontal and vertical alignment settings
		style.setHorizontalAlignment(TextAlignmentType.CENTER);
		style.setVerticalAlignment(TextAlignmentType.CENTER);

		// Apply the style to the cell
		worksheet.getCells().get("B3").setStyle(style);

		// Set the third row height in pixels
		worksheet.getCells().setRowHeightPixel(2, 53);

		// Merge the range of cells (B3:C3)
		worksheet.getCells().merge(2, 1, 1, 2);

		// Save the Excel file
		workbook.save(dataDir + "ApplyGradientFillEffects_out.xlsx");
	}
}
