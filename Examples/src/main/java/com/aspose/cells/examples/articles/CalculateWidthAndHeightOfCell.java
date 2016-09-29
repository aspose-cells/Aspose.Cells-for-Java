package com.aspose.cells.examples.articles;

import com.aspose.cells.Cell;
import com.aspose.cells.Style;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class CalculateWidthAndHeightOfCell {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(CalculateWidthAndHeightOfCell.class) + "articles/";
		// Create workbook object
		Workbook workbook = new Workbook();

		// Access first worksheet
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Access cell B2 and add some value inside it
		Cell cell = worksheet.getCells().get("B2");
		cell.putValue("Welcome to Aspose!");

		// Enlarge its font to size 16
		Style style = cell.getStyle();
		style.getFont().setSize(16);
		cell.setStyle(style);

		// Calculate the width and height of the cell value in unit of pixels
		int widthOfValue = cell.getWidthOfValue();
		int heightOfValue = cell.getHeightOfValue();

		// Print both values
		System.out.println("Width of Cell Value: " + widthOfValue);
		System.out.println("Height of Cell Value: " + heightOfValue);

		// Set the row height and column width to adjust/fit the cell value inside cell
		worksheet.getCells().setColumnWidthPixel(1, widthOfValue);
		worksheet.getCells().setRowHeightPixel(1, heightOfValue);

		// Save the output excel file
		workbook.save(dataDir + "CWAHOfCell_out.xlsx");

	}
}
