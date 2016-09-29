package com.aspose.cells.examples.articles;

import com.aspose.cells.Comment;
import com.aspose.cells.TextAlignmentType;
import com.aspose.cells.TextDirectionType;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class ChangeTextDirectionofComment {
	public static void main(String[] args) throws Exception {

		String dataDir = Utils.getSharedDataDir(ChangeTextDirectionofComment.class) + "articles/";
		// Instantiate a new Workbook
		Workbook wb = new Workbook();
		// Get the first worksheet
		Worksheet sheet = wb.getWorksheets().get(0);

		// Add a comment to A1 cell
		Comment comment = sheet.getComments().get(sheet.getComments().add("A1"));
		// Set its vertical alignment setting
		comment.getCommentShape().setTextVerticalAlignment(TextAlignmentType.CENTER);
		// Set its horizontal alignment setting
		comment.getCommentShape().setTextHorizontalAlignment(TextAlignmentType.RIGHT);
		// Set the Text Direction - Right-to-Left
		comment.getCommentShape().setTextDirection(TextDirectionType.RIGHT_TO_LEFT);
		// Set the Comment note
		comment.setNote("This is my Comment Text. This is test");

		// Save the Excel file
		wb.save(dataDir + "CTDOfComment_out.xlsx");


	}
}
