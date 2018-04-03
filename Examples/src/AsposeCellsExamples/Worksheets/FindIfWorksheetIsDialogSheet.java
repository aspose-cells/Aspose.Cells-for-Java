package AsposeCellsExamples.Worksheets;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class FindIfWorksheetIsDialogSheet { 
	
	static String srcDir = Utils.Get_SourceDirectory();
	static String outDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {

		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());
		
		//Load Excel file containing Dialog Sheet
		Workbook wb = new Workbook(srcDir + "sampleFindIfWorksheetIsDialogSheet.xlsx");

		//Access first worksheet
		Worksheet ws = wb.getWorksheets().get(0);
		 
		//Find if the sheet type is dialog and print the message
		if(ws.getType() == SheetType.DIALOG)
		{ 
		    System.out.println("Worksheet is a Dialog Sheet.");
		}

		// Print the message
		System.out.println("FindIfWorksheetIsDialogSheet executed successfully.");
	}
}
