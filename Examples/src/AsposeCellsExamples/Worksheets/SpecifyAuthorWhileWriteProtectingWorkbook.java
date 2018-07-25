package AsposeCellsExamples.Worksheets;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class SpecifyAuthorWhileWriteProtectingWorkbook { 
	
	static String srcDir = Utils.Get_SourceDirectory();
	static String outDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {

		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());
		
		// Create empty workbook.
		Workbook wb = new Workbook();

		// Write protect workbook with password.
		wb.getSettings().getWriteProtection().setPassword("1234");

		// Specify author while write protecting workbook.
		wb.getSettings().getWriteProtection().setAuthor("SimonAspose");

		// Save the workbook in XLSX format.
		wb.save(outDir + "outputSpecifyAuthorWhileWriteProtectingWorkbook.xlsx");
		
		// Print the message
		System.out.println("SpecifyAuthorWhileWriteProtectingWorkbook executed successfully.");
	}
}
