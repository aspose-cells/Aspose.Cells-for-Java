package AsposeCellsExamples.HTML;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class ExportDocumentWorkbookAndWorksheetPropertiesInHTML { 
	
	static String srcDir = Utils.Get_SourceDirectory();
	static String outDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {

		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());
		
		//Load the sample Excel file
		Workbook workbook = new Workbook(srcDir + "sampleExportDocumentWorkbookAndWorksheetPropertiesInHTML.xlsx");

		//Specify Html Save Options
		HtmlSaveOptions options = new HtmlSaveOptions();

		//We do not want to export document, workbook and worksheet properties
		options.setExportDocumentProperties(false);
		options.setExportWorkbookProperties(false);
		options.setExportWorksheetProperties(false);

		//Export the Excel file to Html with Html Save Options
		workbook.save(outDir + "outputExportDocumentWorkbookAndWorksheetPropertiesInHTML.html", options);

		// Print the message
		System.out.println("ExportDocumentWorkbookAndWorksheetPropertiesInHTML executed successfully.");
	}
}
