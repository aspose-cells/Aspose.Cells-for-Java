package com.aspose.cells.examples.DrawingObjects.comments;

import com.aspose.cells.Comment;
import com.aspose.cells.Font;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class CommentFormatting {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getDataDir(CommentFormatting.class);

		// Instantiating a Workbook object
		Workbook workbook = new Workbook();

		// Adding a new worksheet to the Workbook object
		int sheetIndex = workbook.getWorksheets().add();
		Worksheet worksheet = workbook.getWorksheets().get(sheetIndex);

		// Adding a comment to "F5" cell
		int commentIndex = worksheet.getComments().add("F5");
		Comment comment = worksheet.getComments().get(commentIndex);

		// Setting the font size of a comment to 14
		Font font = comment.getFont();
		font.setSize(14);
		// Setting the font of a comment to bold
		font.setBold(true);

		// Setting the height of the font to 10
		comment.setHeightCM(10);

		// Setting the width of the font to 2
		comment.setWidthCM(2);
		// Saving the Excel file
		workbook.save(dataDir + "book1.xls");
	}
}
