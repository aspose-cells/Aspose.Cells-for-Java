package featurescomparison.workingwithdata.datecells.java;

import java.util.Calendar;

import com.aspose.cells.Cell;
import com.aspose.cells.Cells;
import com.aspose.cells.Style;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

public class AsposeCreateDateCells
{
	public static void main(String[] args) throws Exception
	{
		String dataPath = "src/featurescomparison/workingwithdata/datecells/data/";
		
		//Instantiating a Workbook object
		Workbook workbook = new Workbook();

		//Accessing the added worksheet in the Excel file
		Worksheet worksheet = workbook.getWorksheets().get(0);
		Cells cells = worksheet.getCells();

		//Adding the current system date to "A1" cell
		Cell cell = cells.get("A1");
		cell.setValue(Calendar.getInstance());

		//Setting the display format of the date to number 15 to show date as "d-mmm-yy"
		Style style = cell.getStyle();
		style.setCustom("d-mmm-yy");
		cell.setStyle(style);

		//Saving the modified Excel file in default format
		workbook.save(dataPath + "AsposeDateWorkbook.xls");
		
		System.out.println("Done.");
	}
}
