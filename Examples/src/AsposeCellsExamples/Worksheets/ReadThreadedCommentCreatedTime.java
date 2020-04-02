package AsposeCellsExamples.Worksheets;

import AsposeCellsExamples.Utils;
import com.aspose.cells.ThreadedComment;
import com.aspose.cells.ThreadedCommentCollection;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

public class ReadThreadedCommentCreatedTime {
	public static void main(String[] args) throws Exception {
		// ExStart:1
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ReadThreadedCommentCreatedTime.class) + "Worksheets/";
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
        	System.out.println("Created Time: " + comment.getCreatedTime());
        }
        // ExEnd:1

		System.out.println("ReadThreadedCommentCreatedTime executed successfully.");
	}
}
