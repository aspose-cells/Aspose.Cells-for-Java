package featurescomparison.workingwithworksheets.reordersheetswithinworkbook.java;

import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.WorksheetCollection;

public class AsposeMoveSheetWithinWorkbook
{
	public static void main(String[] args) throws Exception
	{
		String dataPath = "src/featurescomparison/workingwithworksheets/reordersheetswithinworkbook/data/";
		
		//Create a new Workbook.
		Workbook workbook = new Workbook();

		WorksheetCollection worksheets = workbook.getWorksheets();
        Worksheet worksheet1 = worksheets.get(0);
        Worksheet worksheet2 = worksheets.add("Sheet2");
        Worksheet worksheet3 = worksheets.add("Sheet3");
        
		//Move Sheets with in Workbook.
        worksheet2.moveTo(0);
        worksheet1.moveTo(1);
        worksheet3.moveTo(2);

		//Save the excel file.
        workbook.save(dataPath + "AsposeMoveSheet_Out.xls");
		
		System.out.println("Sheet moved successfully."); // Print Message
	}
}
