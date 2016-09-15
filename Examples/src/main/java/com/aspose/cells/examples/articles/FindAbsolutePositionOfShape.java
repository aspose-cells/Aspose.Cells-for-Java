package com.aspose.cells.examples.articles;

import com.aspose.cells.Shape;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class FindAbsolutePositionOfShape {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(FindAbsolutePositionOfShape.class) + "articles/";
		// Load the sample Excel file inside the workbook object
		Workbook workbook = new Workbook("sample.xlsx");

		// Access the first worksheet
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Access the first shape inside the worksheet
		Shape shape = worksheet.getShapes().get(0);

		// Displays the absolute position of the shape
		System.out.println("Absolute Position of this Shape is (" + shape.getLeftToCorner() + " , "
				+ shape.getTopToCorner() + ")");

	}
}
