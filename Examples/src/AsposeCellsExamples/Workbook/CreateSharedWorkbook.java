package AsposeCellsExamples.Workbook;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class CreateSharedWorkbook { 
	
	static String srcDir = Utils.Get_SourceDirectory();
	static String outDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {

		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());
		
		//Create Workbook object
		Workbook wb = new Workbook();

		//Share the Workbook
		wb.getSettings().setShared(true);

		//Save the Shared Workbook
		wb.save(outDir + "outputSharedWorkbook.xlsx");
			
		// Print the message
		System.out.println("CreateSharedWorkbook executed successfully.");
	}
}
