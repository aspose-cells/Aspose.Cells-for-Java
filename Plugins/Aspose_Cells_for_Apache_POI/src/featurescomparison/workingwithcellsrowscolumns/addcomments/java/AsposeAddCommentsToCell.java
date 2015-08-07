package featurescomparison.workingwithcellsrowscolumns.addcomments.java;

import com.aspose.cells.Comment;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

public class AsposeAddCommentsToCell 
{
	public static void main(String[] args) throws Exception 
	{
		String dataPath = "src/featurescomparison/workingwithcellsrowscolumns/addcomments/data/";
		
		//Instantiating a Workbook object
		Workbook workbook = new Workbook();

		Worksheet worksheet = workbook.getWorksheets().get(0);

		//Adding a comment to "F5" cell
		int commentIndex = worksheet.getComments().add("F5");
		Comment comment = worksheet.getComments().get(commentIndex);

		//Setting the comment note
		comment.setNote("Hello Aspose!");

		//Saving the Excel file
		workbook.save(dataPath + "AsposeComments.xls");
		
		System.out.println("Done.");
	}
}
