package AsposeCellsExamples.LoadingSavingConvertingAndManaging;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class ExportCommentsWhileSavingExcelFileToHtml { 

	public static void main(String[] args) throws Exception {

		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());

		String srcDir = Utils.Get_SourceDirectory();
		String outDir = Utils.Get_OutputDirectory();
		
		// Load sample Excel file
		Workbook wb = new Workbook(srcDir + "sampleExportCommentsHTML.xlsx");
		 
		// Export comments - set IsExportComments property to true
		HtmlSaveOptions opts = new HtmlSaveOptions();
		opts.setExportComments(true);
		 
		// Save the Excel file to HTML
		wb.save(outDir + "outputExportCommentsHTML.html", opts);

		// Print the message
		System.out.println("ExportCommentsWhileSavingExcelFileToHtml executed successfully.");
	}
}
