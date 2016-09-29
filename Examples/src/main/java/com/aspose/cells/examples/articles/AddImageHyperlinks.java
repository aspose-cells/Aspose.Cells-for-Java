package com.aspose.cells.examples.articles;

import com.aspose.cells.PlacementType;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class AddImageHyperlinks {
	public static void main(String[] args) throws Exception {

		String dataDir = Utils.getSharedDataDir(AddImageHyperlinks.class) + "articles/";

		// Instantiate a new workbook
		Workbook workbook = new Workbook();

		// Get the first worksheet
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Insert a string value to a cell
		worksheet.getCells().get("C2").setValue("Image Hyperlink");

		// Set the 4th row height
		worksheet.getCells().setRowHeight(3, 100);

		// Set the C column width
		worksheet.getCells().setColumnWidth(2, 21);

		// Add a picture to the C4 cell
		int index = worksheet.getPictures().add(3, 2, 4, 3, dataDir + "aspose-logo.jpg");

		// Get the picture object
		com.aspose.cells.Picture pic = worksheet.getPictures().get(index);

		// Set the placement type
		pic.setPlacement(PlacementType.FREE_FLOATING);

		// Add an image hyperlink
		pic.addHyperlink("http://www.aspose.com/");
		com.aspose.cells.Hyperlink hlink = pic.getHyperlink();

		// Specify the screen tip
		hlink.setScreenTip("Click to go to Aspose site");

		// Save the Excel file
		workbook.save(dataDir + "AIHyperlinks_out.xls");

	}
}
