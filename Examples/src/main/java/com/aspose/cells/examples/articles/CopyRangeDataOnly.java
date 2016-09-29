package com.aspose.cells.examples.articles;

import com.aspose.cells.BackgroundType;
import com.aspose.cells.BorderType;
import com.aspose.cells.CellBorderType;
import com.aspose.cells.Cells;
import com.aspose.cells.Color;
import com.aspose.cells.Range;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Style;
import com.aspose.cells.StyleFlag;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class CopyRangeDataOnly {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(CopyRangeDataOnly.class) + "articles/";
		// Instantiate a new Workbook
		Workbook workbook = new Workbook();

		// Get the first Worksheet Cells
		Cells cells = workbook.getWorksheets().get(0).getCells();

		// Fill some sample data into the cells
		for (int i = 0; i < 50; i++) {
			for (int j = 0; j < 10; j++) {
				cells.get(i, j).putValue(i + "," + j);
			}

		}

		// Create a range (A1:D3).
		Range range = cells.createRange("A1", "D3");

		// Create a style object.
		Style style = workbook.createStyle();

		// Specify the font attribute.
		style.getFont().setName("Calibri");

		// Specify the shading color.
		style.setForegroundColor(Color.getYellow());
		style.setPattern(BackgroundType.SOLID);

		// Specify the border attributes.
		style.getBorders().getByBorderType(BorderType.TOP_BORDER).setLineStyle(CellBorderType.THIN);
		style.getBorders().getByBorderType(BorderType.TOP_BORDER).setColor(Color.getBlue());
		style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setLineStyle(CellBorderType.THIN);
		style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER).setColor(Color.getBlue());
		style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setLineStyle(CellBorderType.THIN);
		style.getBorders().getByBorderType(BorderType.LEFT_BORDER).setColor(Color.getBlue());
		style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setLineStyle(CellBorderType.THIN);
		style.getBorders().getByBorderType(BorderType.RIGHT_BORDER).setColor(Color.getBlue());

		// Create the styleflag object.
		StyleFlag flag = new StyleFlag();

		// Implement font attribute
		flag.setFontName(true);

		// Implement the shading / fill color.
		flag.setCellShading(true);

		// Implment border attributes.
		flag.setBorders(true);

		// Set the Range style.
		range.applyStyle(style, flag);

		// Create a second range (L9:O11)
		Range range2 = cells.createRange("L9", "O11");

		// Copy the range data only.
		range2.copyData(range);

		// Save the Excel file.
		workbook.save(dataDir + "CopyRangeDataOnly_out.xlsx", SaveFormat.XLSX);


	}
}
