package com.aspose.cells.examples.DrawingObjects.comments;

import com.aspose.cells.Comment;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;
import com.aspose.cells.examples.DrawingObjects.NonPrimitiveShape;

public class AddingComment {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(AddingComment.class) + "DrawingObjects/";
		// Instantiating a Workbook object
		Workbook workbook = new Workbook();

		// Adding a new worksheet to the Workbook object
		int sheetIndex = workbook.getWorksheets().add();
		Worksheet worksheet = workbook.getWorksheets().get(sheetIndex);

		// Adding a comment to "F5" cell
		int commentIndex = worksheet.getComments().add("F5");
		Comment comment = worksheet.getComments().get(commentIndex);

		// Setting the comment note
		comment.setNote("Hello Aspose!");

		// Saving the Excel file
		workbook.save(dataDir + "AComment-out.xls");
	}
}
