package com.aspose.cells.examples.DrawingObjects;

import com.aspose.cells.Picture;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class AbsolutePositioning {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(AbsolutePositioning.class) + "DrawingObjects/";

		// Instantiating a Workbook object
		Workbook workbook = new Workbook();

		// Obtaining the reference of the newly added worksheet.
		int sheetIndex = workbook.getWorksheets().add();
		Worksheet worksheet = workbook.getWorksheets().get(sheetIndex);

		// Adding a picture at the location of a cell whose row and column indices are 5 in the worksheet. It is "F6" cell
		int pictureIndex = worksheet.getPictures().add(5, 5, dataDir + "logo.jpg");
		Picture picture = worksheet.getPictures().get(pictureIndex);

		// Positioning the picture proportional to row height and colum width
		picture.setUpperDeltaX(200);
		picture.setUpperDeltaY(200);

		// Saving the Excel file
		workbook.save(dataDir + "APositioning_out.xls");
	}
}
