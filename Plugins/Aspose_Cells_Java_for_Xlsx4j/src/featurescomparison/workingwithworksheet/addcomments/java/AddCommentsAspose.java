package featurescomparison.workingwithworksheet.addcomments.java;

import com.aspose.cells.Comment;
import com.aspose.cells.Font;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

/**
 * @author Shoaib Khan
 */
public class AddCommentsAspose
{
	public static void main(String[] args) throws Exception
	{
		String dataPath = "src/featurescomparison/workingwithworksheet/addcomments/data/";
		
		// Instantiating a Workbook object
		Workbook workbook = new Workbook();

		// Adding a new worksheet to the Workbook object
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Adding a comment to cell
		int commentIndex = worksheet.getComments().add("A1");
		Comment comment = worksheet.getComments().get(commentIndex);

		// Setting the comment note
		comment.setNote("Hello Aspose!");

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
		workbook.save(dataPath + "AddComments-Aspose.xlsx", SaveFormat.XLSX);

		// Print Message
		System.out.println("Comment added successfully.");
	}
}
