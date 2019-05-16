package AsposeCellsExamples.Worksheets;

import com.aspose.cells.ThreadedComment;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import AsposeCellsExamples.Utils;
public class EditThreadedComments {
	
	public static void main(String[] args) throws Exception {
		// ExStart:1
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(CountNumberOfCells.class) + "Worksheets/";
		// Instantiating a Workbook object
		Workbook workbook = new Workbook(dataDir + "ThreadedCommentsSample.xlsx");

		//Access first worksheet
        Worksheet worksheet = workbook.getWorksheets().get(0);

        // Get Threaded Comment
        ThreadedComment comment = worksheet.getComments().getThreadedComments("A1").get(0);
        comment.setNotes("Updated Comment");

        workbook.save(dataDir + "EditThreadedComments.xlsx");
        // ExEnd:1

		System.out.println("EditThreadedComments executed successfully.");
	}
}