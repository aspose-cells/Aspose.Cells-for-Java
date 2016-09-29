package com.aspose.cells.examples.articles;

import com.aspose.cells.Color;
import com.aspose.cells.ThemeColorType;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class GetSetThemeColors {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(GetSetThemeColors.class) + "articles/";
		// Instantiate Workbook object
		// Open an exiting excel file
		Workbook workbook = new Workbook(dataDir + "book1.xlsx");

		// Get the Background1 theme color
		Color c = workbook.getThemeColor(ThemeColorType.BACKGROUND_1);
		// Print the color
		System.out.println("theme color Background1: " + c);

		// Get the Accent2 theme color
		c = workbook.getThemeColor(ThemeColorType.ACCENT_1);
		// Print the color
		System.out.println("theme color Accent2: " + c);

		// Change the Background1 theme color
		workbook.setThemeColor(ThemeColorType.BACKGROUND_1, Color.getRed());
		// Get the updated Background1 theme color
		c = workbook.getThemeColor(ThemeColorType.BACKGROUND_1);
		// Print the updated color for confirmation
		System.out.println("theme color Background1 changed to: " + c);

		// Change the Accent2 theme color
		workbook.setThemeColor(ThemeColorType.ACCENT_1, Color.getBlue());
		// Get the updated Accent2 theme color
		c = workbook.getThemeColor(ThemeColorType.ACCENT_1);
		// Print the updated color for confirmation
		System.out.println("theme color Accent2 changed to: " + c);

		// Save the updated file
		workbook.save(dataDir + "GetSetThemeColors_out.xlsx");

	}
}
