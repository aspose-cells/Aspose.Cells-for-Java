package com.aspose.cells.examples.articles;

import com.aspose.cells.Color;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class CustomizingThemes {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(CustomizingThemes.class) + "articles/";
		// Define Color array (of 12 colors) for the Theme
		Color[] carr = new Color[12];
		carr[0] = Color.getAntiqueWhite(); // Background1
		carr[1] = Color.getBrown(); // Text1
		carr[2] = Color.getAliceBlue(); // Background2
		carr[3] = Color.getYellow(); // Text2
		carr[4] = Color.getYellowGreen(); // Accent1
		carr[5] = Color.getRed(); // Accent2
		carr[6] = Color.getPink(); // Accent3
		carr[7] = Color.getPurple(); // Accent4
		carr[8] = Color.getPaleGreen(); // Accent5
		carr[9] = Color.getOrange(); // Accent6
		carr[10] = Color.getGreen(); // Hyperlink
		carr[11] = Color.getGray(); // Followed Hyperlink

		// Instantiate a Workbook
		// Open the spreadsheet file
		Workbook workbook = new Workbook(dataDir + "book1.xlsx");

		// Set the custom theme with specified colors
		workbook.customTheme("CustomeTheme1", carr);

		// Save as the excel file
		workbook.save(dataDir + "CustomizingThemes_out.xlsx");

	}
}
