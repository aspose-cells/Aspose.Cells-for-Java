package com.aspose.cells.examples.articles;

import com.aspose.cells.Picture;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class InsertLinkedPicturefromWebAddress {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(InsertLinkedPicturefromWebAddress.class) + "articles/";
		// Instantiate a new Workbook.
		Workbook workbook = new Workbook();
		// Insert a linked picture (from Web Address) to B2 Cell.
		Picture pic = (Picture) workbook.getWorksheets().get(0).getShapes().addLinkedPicture(1, 1, 100, 100,
				"http://www.aspose.com/Images/aspose-logo.jpg");
		// Set the source of the inserted image.
		pic.setSourceFullName("http://www.aspose.com/images/aspose-logo.gif");

		// Set the height and width of the inserted image.
		pic.setHeightInch(1.04);
		pic.setWidthInch(2.6);

		// Save the Excel file.
		workbook.save(dataDir + "ILPfromWebAddress_out.xlsx");

	}
}
