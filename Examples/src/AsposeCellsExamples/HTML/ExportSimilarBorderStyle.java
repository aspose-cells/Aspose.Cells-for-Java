package AsposeCellsExamples.HTML;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class ExportSimilarBorderStyle {
	
	static String srcDir = Utils.Get_SourceDirectory();
	static String outDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {

		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());
		
		//Load the sample Excel file
		Workbook wb = new Workbook(srcDir + "sampleExportSimilarBorderStyle.xlsx");
		 
		//Specify Html Save Options - Export Similar Border Style
		HtmlSaveOptions opts = new HtmlSaveOptions();
		opts.setExportSimilarBorderStyle(true);
		 
		//Save the workbook in Html format with specified Html Save Options
		wb.save(outDir + "outputExportSimilarBorderStyle.html", opts);


		// Print the message
		System.out.println("ExportSimilarBorderStyle executed successfully.");
	}
}
