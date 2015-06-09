package featurescomparison.workingwithcellsrowscolumns.freezepanes.java;

import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.WorksheetCollection;

public class AsposeFreezePanes
{
	public static void main(String[] args) throws Exception
	{
		String dataPath = "src/featurescomparison/workingwithcellsrowscolumns/freezepanes/data/";
		
		//Instantiating a Excel object by excel file path
        Workbook workbook = new Workbook();

        //Accessing the first worksheet in the Excel file
        WorksheetCollection worksheets = workbook.getWorksheets();

        Worksheet worksheet1 = worksheets.get(0);
        Worksheet worksheet2 = worksheets.add("Sheet2");       

        //Applying freeze panes settings
        worksheet1.freezePanes(0,2,0,2); // Freezing Columns
        worksheet2.freezePanes(2,0,2,0); // Freezing Rows     

        //Saving the modified Excel file in default format
        workbook.save(dataPath + "workbook_Aspose_Out.xls");
        
        //Print Message
        System.out.println("Panes freeze successfull.");
	}
}