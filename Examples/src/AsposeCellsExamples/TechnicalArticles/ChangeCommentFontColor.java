package AsposeCellsExamples.TechnicalArticles;

import AsposeCellsExamples.Utils;
import com.aspose.cells.*;

public class ChangeCommentFontColor {
	public static void main(String[] args) throws Exception {
		// ExStart:1
		String dataDir = Utils.Get_OutputDirectory();
		// Instantiate a new Workbook
		Workbook workbook = new Workbook();
		// Get the first worksheet
		Worksheet worksheet = workbook.getWorksheets().get(0);

		//Add some text in cell A1
		worksheet.getCells().get("A1").putValue("Here");

		// Add a comment to A1 cell
		Comment comment = worksheet.getComments().get(worksheet.getComments().add("A1"));
		// Set its vertical alignment setting
		comment.getCommentShape().setTextVerticalAlignment(TextAlignmentType.CENTER);
		// Set the Comment note
		comment.setNote("This is my Comment Text. This is test");

		Shape shape = worksheet.getComments().get("A1").getCommentShape();

		shape.getFill().getSolidFill().setColor(Color.getBlack());
		Font font = shape.getFont();
		font.setColor(Color.getWhite());
		StyleFlag styleFlag = new StyleFlag();
		styleFlag.setFontColor(true);
		shape.getTextBody().format(0, shape.getText().length(), font, styleFlag);

		// Save the Excel file
		workbook.save(dataDir + "outputChangeCommentFontColor.xlsx");
		// ExEnd:1

		System.out.println("ChangeCommentFontColor executed successfully.\r\n");
	}
}
