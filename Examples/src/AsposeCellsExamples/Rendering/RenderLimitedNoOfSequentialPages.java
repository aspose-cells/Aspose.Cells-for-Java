package AsposeCellsExamples.Rendering;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class RenderLimitedNoOfSequentialPages { 
	
	static String srcDir = Utils.Get_SourceDirectory();
	static String outDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {

		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());
		
		//Load the sample Excel file
		Workbook wb = new Workbook(srcDir + "sampleImageOrPrintOptions_PageIndexPageCount.xlsx");
		  
		//Access the first worksheet
		Worksheet ws = wb.getWorksheets().get(0);
		  
		//Specify image or print options
		//We want to print pages 4, 5, 6, 7
		ImageOrPrintOptions opts = new ImageOrPrintOptions();
		opts.setPageIndex(3);
		opts.setPageCount(4);
		opts.setImageFormat(ImageFormat.getPng());
		  
		//Create sheet render object
		SheetRender sr = new SheetRender(ws, opts);
		  
		//Print all the pages as images
		for (int i = opts.getPageIndex(); i < sr.getPageCount(); i++)
		{
		    sr.toImage(i, outDir + "outputImage-" + (i+1) + ".png");
		}
		
		// Print the message
		System.out.println("RenderLimitedNoOfSequentialPages executed successfully.");
	}
}
