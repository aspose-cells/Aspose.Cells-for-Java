package featurescomparison.workingwithformattingfeatures.workingwithborders.java;

import com.aspose.cells.BorderType;
import com.aspose.cells.Cell;
import com.aspose.cells.CellBorderType;
import com.aspose.cells.Cells;
import com.aspose.cells.Color;
import com.aspose.cells.Style;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

public class AsposeBorders
{
	public static void main(String[] args) throws Exception
	{
		String dataPath = "src/featurescomparison/workingwithformattingfeatures/workingwithborders/data/";
		
		//Instantiating a Workbook object
		Workbook workbook = new Workbook();

		//Accessing the worksheet in the Excel file
		Worksheet worksheet = workbook.getWorksheets().get(0);
		Cells cells = worksheet.getCells();

		//Accessing the "A1" cell from the worksheet      
		Cell cell = cells.get("B2");

		//Adding some value to the "A1" cell
		cell.setValue("Visit Aspose @ www.aspose.com!");
		Style style = cell.getStyle();

		//Setting the line of the top border
		style.setBorder(BorderType.TOP_BORDER,CellBorderType.THICK,Color.getBlack());

		//Setting the line of the bottom border
		style.setBorder(BorderType.BOTTOM_BORDER,CellBorderType.THICK,Color.getBlack());

		//Setting the line of the left border
		style.setBorder(BorderType.LEFT_BORDER,CellBorderType.THICK,Color.getBlack());

		//Setting the line of the right border
		style.setBorder(BorderType.RIGHT_BORDER,CellBorderType.THICK,Color.getBlack());

		//Saving the modified style to the "A1" cell.
		cell.setStyle(style);

		//Saving the Excel file
		workbook.save(dataPath + "AsposeBorders_Out.xls");

		System.out.println("Aspose Borders Created.");
	}
}
