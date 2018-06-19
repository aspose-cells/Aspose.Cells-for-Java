package AsposeCellsExamples.LoadingSavingConvertingAndManaging;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class SaveWorkbookToStrictOpenXMLSpreadsheetFormat { 
	
	static String outDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {

		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());
		
		// Create workbook.
		Workbook wb = new Workbook();
		 
		// Specify - Strict Open XML Spreadsheet - Format.
		wb.getSettings().setCompliance(OoxmlCompliance.ISO_29500_2008_STRICT);
		 
		// Add message in cell B4 of first worksheet.
		Cell b4 = wb.getWorksheets().get(0).getCells().get("B4");
		b4.putValue("This Excel file has Strict Open XML Spreadsheet format.");
		 
		// Save to output Excel file.
		wb.save(outDir + "outputSaveWorkbookToStrictOpenXMLSpreadsheetFormat.xlsx", SaveFormat.XLSX);

		// Print the message
		System.out.println("SaveWorkbookToStrictOpenXMLSpreadsheetFormat executed successfully.");
	}
}
