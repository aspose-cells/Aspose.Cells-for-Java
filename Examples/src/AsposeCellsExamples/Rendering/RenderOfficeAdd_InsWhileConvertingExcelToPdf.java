package AsposeCellsExamples.Rendering;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class RenderOfficeAdd_InsWhileConvertingExcelToPdf {
	
	static String srcDir = Utils.Get_SourceDirectory();
	static String outDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {

		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());
		
		//Load the sample Excel file containing Office Add-Ins
		Workbook wb = new Workbook(srcDir + "sampleRenderOfficeAdd-Ins.xlsx");

		//Save it to Pdf format
		wb.save(outDir + "output-"  + CellsHelper.getVersion() + ".pdf");

		// Print the message
		System.out.println("RenderOfficeAdd_InsWhileConvertingExcelToPdf executed successfully.");
	}
}
