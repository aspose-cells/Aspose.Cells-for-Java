package AsposeCellsExamples.Rendering;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class OutputBlankPageWhenThereIsNothingToPrint { 

	public static void main(String[] args) throws Exception {

		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());
		// ExStart:1

		String outDir = Utils.Get_OutputDirectory();
		
		// Create workbook
		Workbook wb = new Workbook();
		 
		// Access first worksheet - it is empty sheet
		Worksheet ws = wb.getWorksheets().get(0);
		 
		// Specify image or print options
		// Since the sheet is blank, we will set
		// OutputBlankPageWhenNothingToPrint to true
		// So that empty page gets printed
		ImageOrPrintOptions opts = new ImageOrPrintOptions();
		opts.setImageType(ImageType.PNG);
		opts.setOutputBlankPageWhenNothingToPrint(true);
		 
		// Render empty sheet to png image
		SheetRender sr = new SheetRender(ws, opts);
		sr.toImage(0, outDir + "OutputBlankPageWhenNothingToPrint.png");
		// ExEnd:1

		// Print the message
		System.out.println("OutputBlankPageWhenThereIsNothingToPrint executed successfully.");
	}
}
