package com.aspose.cells.examples.articles;

import com.aspose.cells.BackgroundType;
import com.aspose.cells.Cell;
import com.aspose.cells.Cells;
import com.aspose.cells.Font;
import com.aspose.cells.Style;
import com.aspose.cells.ThemeColor;
import com.aspose.cells.ThemeColorType;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class UseThemeColors {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(UseThemeColors.class) + "articles/";
		// Instantiate a Workbook
		Workbook workbook = new Workbook();
		// Get cells collection in the first (default) worksheet
		Cells cells = workbook.getWorksheets().get(0).getCells();
		// Get the D3 cell
		Cell c = cells.get("D3");

		// Get the style of the cell
		Style s = c.getStyle();
		// Set background color for the cell from the default theme Accent2 color
		s.setBackgroundThemeColor(new ThemeColor(ThemeColorType.ACCENT_2, 0.5));
		// Set the pattern type
		s.setPattern(BackgroundType.SOLID);
		// Get the font for the style
		Font f = s.getFont();
		// Set the theme color
		f.setThemeColor(new ThemeColor(ThemeColorType.ACCENT_4, 0.1));

		// Apply style
		c.setStyle(s);

		// Put a value
		c.putValue("Testing");

		// Save the excel file
		workbook.save(dataDir + "UseThemeColors_out.xlsx");

	}
}
