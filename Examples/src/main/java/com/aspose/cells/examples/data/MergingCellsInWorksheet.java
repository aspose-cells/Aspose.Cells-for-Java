package com.aspose.cells.examples.data;

import com.aspose.cells.BackgroundType;
import com.aspose.cells.Cells;
import com.aspose.cells.Color;
import com.aspose.cells.Font;
import com.aspose.cells.Style;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class MergingCellsInWorksheet {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(MergingCellsInWorksheet.class) + "data/";

		// Create a Workbook.
		Workbook wbk = new Workbook();

		// Create a Worksheet and get the first sheet.
		Worksheet worksheet = wbk.getWorksheets().get(0);

		// Create a Cells object to fetch all the cells.
		Cells cells = worksheet.getCells();

		// Merge some Cells (C6:E7) into a single C6 Cell.
		cells.merge(5, 2, 2, 3);

		// Input data into C6 Cell.
		worksheet.getCells().get(5, 2).setValue("This is my value");

		// Create a Style object to fetch the Style of C6 Cell.
		Style style = worksheet.getCells().get(5, 2).getStyle();

		// Create a Font object
		Font font = style.getFont();

		// Set the name.
		font.setName("Times New Roman");

		// Set the font size.
		font.setSize(18);

		// Set the font color
		font.setColor(Color.getBlue());

		// Bold the text
		font.setBold(true);

		// Make it italic
		font.setItalic(true);

		// Set the backgrond color of C6 Cell to Red
		style.setForegroundColor(Color.getRed());
		style.setPattern(BackgroundType.SOLID);

		// Apply the Style to C6 Cell.
		cells.get(5, 2).setStyle(style);

		// Save the Workbook.
		wbk.save(dataDir + "mergingcells_out.xls");
		wbk.save(dataDir + "mergingcells_out.xlsx");
		wbk.save(dataDir + "mergingcells_out.ods");

		// Print message
		System.out.println("Process completed successfully");

	}
}
