package featurescomparison.workingwithcellsrowscolumns.splitpanes.java;

import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;

public class AsposeSplitPanes 
{
	public static void main(String[] args) throws Exception 
	{
		String dataPath = "src/featurescomparison/workingwithcellsrowscolumns/splitpanes/data/";
		
		//Instantiate a new workbook / Open a template file
		Workbook book = new Workbook(dataPath + "workbook.xls");

		//Set the active cell
		book.getWorksheets().get(0).setActiveCell("A20");

		//Split the worksheet window
		book.getWorksheets().get(0).split();

		//Save the Excel file
		book.save(dataPath + "AsposeSplitPanes.xls", SaveFormat.EXCEL_97_TO_2003);
		
		System.out.println("Done.");
	}
}
