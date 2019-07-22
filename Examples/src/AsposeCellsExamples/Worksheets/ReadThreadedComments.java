package AsposeCellsExamples.Worksheets;

import com.aspose.cells.ThreadedComment;
import com.aspose.cells.ThreadedCommentCollection;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import AsposeCellsExamples.Utils;
public class ReadThreadedComments {
	
	public static void main(String[] args) throws Exception {
		// ExStart:1
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ReadThreadedComments.class) + "Worksheets/";
		// Instantiating a Workbook object
		Workbook workbook = new Workbook(dataDir + "ThreadedCommentsSample.xlsx");

		//Access first worksheet
        Worksheet worksheet = workbook.getWorksheets().get(0);

        // Get Threaded Comments
        ThreadedCommentCollection threadedComments = worksheet.getComments().getThreadedComments("A1");

        for (Object obj : threadedComments)
        {
        	ThreadedComment comment = (ThreadedComment) obj;
        	System.out.println("Comment: " + comment.getNotes());
        	System.out.println("Author: " + comment.getAuthor().getName());
        }
        // ExEnd:1

		System.out.println("ReadThreadedComments executed successfully.");
	}
}
