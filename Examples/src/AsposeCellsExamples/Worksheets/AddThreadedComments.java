package AsposeCellsExamples.Worksheets;

import com.aspose.cells.ThreadedCommentAuthor;
import com.aspose.cells.Workbook;
import AsposeCellsExamples.Utils;
public class AddThreadedComments {
	
	public static void main(String[] args) throws Exception {
		// ExStart:1
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(CountNumberOfCells.class) + "Worksheets/";
		// Instantiating a Workbook object
		Workbook workbook = new Workbook();

		// Add Author
		int authorIndex = workbook.getWorksheets().getThreadedCommentAuthors().add("Aspose Test", "", "");
		ThreadedCommentAuthor author = workbook.getWorksheets().getThreadedCommentAuthors().get(authorIndex);
		
		// Add Threaded Comment
		workbook.getWorksheets().get(0).getComments().addThreadedComment("A1", "Test Threaded Comment", author);

        workbook.save(dataDir + "AddThreadedComments_out.xlsx");
        // ExEnd:1

		System.out.println("AddThreadedComments executed successfully.");
	}
}
