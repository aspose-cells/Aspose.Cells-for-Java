package AsposeCellsExamples.Rendering;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class IgnoreErrorsWhileRenderingExcelToPdf {
	
	static String srcDir = Utils.Get_SourceDirectory();
	static String outDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {

		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());

		//Load the Sample Workbook that throws Error on Excel2Pdf conversion
		Workbook wb = new Workbook(srcDir + "sampleErrorExcel2Pdf.xlsx");
		 
		//Specify Pdf Save Options - Ignore Error
		PdfSaveOptions opts = new PdfSaveOptions();
		opts.setIgnoreError(true);
		 
		//Save the Workbook in Pdf with Pdf Save Options
		wb.save(outDir + "outputErrorExcel2Pdf.pdf", opts);

		// Print the message
		System.out.println("IgnoreErrorsWhileRenderingExcelToPdf executed successfully.");
	}
}
