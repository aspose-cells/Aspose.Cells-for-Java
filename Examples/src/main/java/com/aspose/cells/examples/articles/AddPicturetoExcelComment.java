package com.aspose.cells.examples.articles;

import java.io.FileInputStream;

import com.aspose.cells.Comment;
import com.aspose.cells.CommentCollection;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class AddPicturetoExcelComment {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(AddPicturetoExcelComment.class) + "articles/";
		// Instantiate a Workbook
		Workbook workbook = new Workbook();

		// Get a reference of comments collection with the first sheet
		CommentCollection comments = workbook.getWorksheets().get(0).getComments();

		// Add a comment to cell A1
		int commentIndex = comments.add(0, 0);
		Comment comment = comments.get(commentIndex);
		comment.setNote("First note.");
		comment.getFont().setName("Times New Roman");

		// Load/Read an image into stream
		String logo_url = dataDir + "school.jpg";

		// Creating the instance of the FileInputStream object to open the logo/picture in the stream
		FileInputStream inFile = new FileInputStream(logo_url);

		// Setting the logo/picture
		byte[] picData = new byte[inFile.available()];
		inFile.read(picData);

		// Set image data to the shape associated with the comment
		comment.getCommentShape().getFillFormat().setImageData(picData);

		// Save the workbook
		workbook.save(dataDir + "APToExcelComment_out.xlsx");

	}
}
