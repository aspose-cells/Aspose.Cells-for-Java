package AsposeCellsExamples.Workbook;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class PasswordProtectOrUnprotectSharedWorkbook { 
	
	static String srcDir = Utils.Get_SourceDirectory();
	static String outDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {

		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());
			
		//Create empty Excel file
		Workbook wb = new Workbook();

		//Protect the Shared Workbook with Password
		wb.protectSharedWorkbook("1234");

		//Uncomment this line to Unprotect the Shared Workbook
		//wb.unprotectSharedWorkbook("1234");

		//Save the output Excel file
		wb.save(outDir + "outputProtectSharedWorkbook.xlsx");
		
		// Print the message
		System.out.println("PasswordProtectOrUnprotectSharedWorkbook executed successfully.");
	}
}
