package AsposeCellsExamples.Slicers;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class RenderingSlicer { 
	
	static String srcDir = Utils.Get_SourceDirectory();
	static String outDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {

		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());
		
		// Load sample Excel file containing slicer.
		Workbook wb = new Workbook(srcDir + "sampleRenderingSlicer.xlsx");

		// Access first worksheet.
		Worksheet ws = wb.getWorksheets().get(0);

		// Set the print area because we want to render slicer only.
		ws.getPageSetup().setPrintArea("B15:E25");

		// Specify image or print options, set one page per sheet and only area to true.
		ImageOrPrintOptions imgOpts = new ImageOrPrintOptions();
		imgOpts.setHorizontalResolution(200);
		imgOpts.setVerticalResolution(200);
		imgOpts.setImageType(com.aspose.cells.ImageType.PNG);
		imgOpts.setOnePagePerSheet(true);
		imgOpts.setOnlyArea(true);

		// Create sheet render object and render worksheet to image.
		SheetRender sr = new SheetRender(ws, imgOpts);
		sr.toImage(0, outDir + "outputRenderingSlicer.png"); 
		
		// Print the message
		System.out.println("RenderingSlicer executed successfully.");
	}
}
