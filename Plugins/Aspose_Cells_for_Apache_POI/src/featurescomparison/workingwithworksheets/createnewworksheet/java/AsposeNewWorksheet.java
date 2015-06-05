package featurescomparison.workingwithworksheets.createnewworksheet.java;

import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.WorksheetCollection;

public class AsposeNewWorksheet
{
	public static void main(String[] args) throws Exception
	{
		String dataPath = "src/featurescomparison/workingwithworksheets/createnewworksheet/data/";
		
        //Instantiating a Workbook object
        Workbook workbook = new Workbook();

		//Adding a new worksheet to the Workbook object
		WorksheetCollection worksheets = workbook.getWorksheets();
		Worksheet worksheet = worksheets.add("My Worksheet");

		//Saving the Excel file
        workbook.save(dataPath + "newWorksheet_Aspose_Out.xls");
        
        //Print Message
        System.out.println("Sheet added successfully.");
	}
}
