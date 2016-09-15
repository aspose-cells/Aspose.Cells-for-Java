package com.aspose.cells.examples.articles;

import com.aspose.cells.Border;
import com.aspose.cells.BorderType;
import com.aspose.cells.Cell;
import com.aspose.cells.Style;
import com.aspose.cells.ThemeColorType;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class ExtractThemeData {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ExtractThemeData.class) + "articles/";
		// Create workbook object
		Workbook workbook = new Workbook(dataDir + "TestBook.xlsx");

		// Extract theme name applied to this workbook
		System.out.println(workbook.getTheme());

		// Access first worksheet
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Access cell A1
		Cell cell = worksheet.getCells().get("A1");

		// Get the style object
		Style style = cell.getStyle();

		// Extract theme color applied to this cell
		System.out.println(style.getForegroundThemeColor().getColorType() == ThemeColorType.ACCENT_2);

		// Extract theme color applied to the bottom border of the cell
		Border bot = style.getBorders().getByBorderType(BorderType.BOTTOM_BORDER);
		System.out.println(bot.getThemeColor().getColorType() == ThemeColorType.ACCENT_1);

	}
}
