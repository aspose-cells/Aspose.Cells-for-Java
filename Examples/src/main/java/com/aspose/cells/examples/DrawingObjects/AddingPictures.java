package com.aspose.cells.examples.DrawingObjects;

import com.aspose.cells.Picture;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.WorksheetCollection;
import com.aspose.cells.examples.Utils;

public class AddingPictures {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(AddingPictures.class) + "DrawingObjects/";

		// Instantiating a Workbook object
		Workbook workbook = new Workbook();

		WorksheetCollection worksheets = workbook.getWorksheets();

		// Obtaining the reference of first worksheet
		Worksheet sheet = worksheets.get(0);

		// Adding a picture at the location of a cell whose row and column indices are 5 in the worksheet. It is "F6" cell
		int pictureIndex = sheet.getPictures().add(5, 5, dataDir + "logo.jpg");
		Picture picture = sheet.getPictures().get(pictureIndex);

		// Saving the Excel file
		workbook.save(dataDir + "AddingPictures_out.xls");
	}
}
