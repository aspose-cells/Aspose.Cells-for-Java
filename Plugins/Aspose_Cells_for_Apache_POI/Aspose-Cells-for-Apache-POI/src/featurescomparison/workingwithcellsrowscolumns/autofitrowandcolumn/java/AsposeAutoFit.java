package featurescomparison.workingwithcellsrowscolumns.autofitrowandcolumn.java;

import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

public class AsposeAutoFit
{
	public static void main(String[] args) throws Exception
	{
		String dataPath = "src/featurescomparison/workingwithcellsrowscolumns/autofitrowandcolumn/data/";
		
        //Instantiating a Workbook object
        Workbook workbook = new Workbook(dataPath + "workbook.xls");

        //Accessing the first worksheet in the Excel file
        Worksheet worksheet = workbook.getWorksheets().get(0);

        worksheet.autoFitRow(1); //Auto-fitting the 2nd row of the worksheet
        worksheet.autoFitColumn(0); //Auto-fitting the 1st column of the worksheet

        //Saving the modified Excel file in default (that is Excel 2003) format
        workbook.save(dataPath + "AutoFit_Aspose_Out.xls");
        
        //Print message
        System.out.println("Row and Column auto fit successfully.");

	}
}
