package AsposeCellsExamples.Worksheets;

import com.aspose.cells.CommentCollection;
import com.aspose.cells.ThreadedCommentAuthor;
import com.aspose.cells.ThreadedCommentAuthorCollection;
import com.aspose.cells.ThreadedCommentCollection;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import AsposeCellsExamples.Utils;
public class RemoveThreadedComments {
	
	public static void main(String[] args) throws Exception {
		// ExStart:1
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(RemoveThreadedComments.class) + "Worksheets/";
		// Instantiating a Workbook object
		Workbook workbook = new Workbook(dataDir + "ThreadedCommentsSample.xlsx");

		//Access first worksheet
        Worksheet worksheet = workbook.getWorksheets().get(0);

        CommentCollection comments = worksheet.getComments();
        ThreadedCommentCollection threadedComments = worksheet.getComments().getThreadedComments("A1");
        ThreadedCommentAuthor author = threadedComments.get(0).getAuthor();
        
        comments.removeAt("I4");
        
        ThreadedCommentAuthorCollection authors = workbook.getWorksheets().getThreadedCommentAuthors();

        authors.removeAt(authors.indexOf(author));
        workbook.save(dataDir + "ThreadedCommentsSample_Out.xlsx");
        // ExEnd:1

		System.out.println("RemoveThreadedComments executed successfully.");
	}
}
