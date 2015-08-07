package featurescomparison.workingwithformattingfeatures.cellalignment.java;

import com.aspose.cells.Cell;
import com.aspose.cells.Cells;
import com.aspose.cells.Style;
import com.aspose.cells.TextAlignmentType;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

public class AsposeCellAlignment 
{
	public static void main(String[] args) throws Exception 
	{
		String dataPath = "src/featurescomparison/workingwithformattingfeatures/cellalignment/data/";
		
		//Instantiating a Workbook object
		Workbook workbook = new Workbook();

		//Accessing the added worksheet in the Excel file
		int sheetIndex = workbook.getWorksheets().add();
		Worksheet worksheet = workbook.getWorksheets().get(sheetIndex);
		Cells cells = worksheet.getCells();

		//Adding the current system date to "A1" cell
		Cell cell = cells.get("A1");
		Style style = cell.getStyle();

		//Adding some value to the "A1" cell
		cell.setValue("Visit Aspose!");

		//Setting the horizontal alignment of the text in the "A1" cell
		style.setHorizontalAlignment(TextAlignmentType.CENTER);

		//Saved style
		cell.setStyle(style);

		//Saving the modified Excel file in default format
		workbook.save(dataPath + "AsposeCellsAlignment.xls");
		
		System.out.println("Done.");
	}
}
