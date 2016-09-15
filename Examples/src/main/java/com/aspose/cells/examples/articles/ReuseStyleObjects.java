package com.aspose.cells.examples.articles;

import com.aspose.cells.Cells;
import com.aspose.cells.Color;
import com.aspose.cells.Font;
import com.aspose.cells.Style;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class ReuseStyleObjects {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ReuseStyleObjects.class) + "articles/";
		// Create an instance of Workbook & load an existing spreadsheet
		Workbook workbook = new Workbook(dataDir + "Book1.xls");

		// Retrieve the Cell Collection of the first Worksheet
		Cells cells = workbook.getWorksheets().get(0).getCells();

		// Create an instance of Style and add it to the pool of styles
		Style styleObject = workbook.createStyle();

		// Retrieve the Font object of newly created style
		Font font = styleObject.getFont();

		// Set the font color to Red
		font.setColor(Color.getRed());

		// Set the newly created style on two different cells
		cells.get("A1").setStyle(styleObject);
		cells.get("A2").setStyle(styleObject);

	}
}
