package AsposeCellsExamples.Worksheets.PageSetupFeatures;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class ImplementCustomPaperSizeOfWorksheetForRendering {
	public static void main(String[] args) throws Exception {
		String srcDir = Utils.Get_SourceDirectory();
		String outDir = Utils.Get_OutputDirectory();

		//Create workbook object
		Workbook wb = new Workbook();
		  
		//Access first worksheet
		Worksheet ws = wb.getWorksheets().get(0);
		  
		//Set custom paper size in unit of inches
		ws.getPageSetup().customPaperSize(6, 4);
		  
		//Access cell B4
		Cell b4 = ws.getCells().get("B4");
		  
		//Add the message in cell B4
		b4.putValue("Pdf Page Dimensions: 6.00 x 4.00 in");
		 
		//Save the workbook in pdf format
		wb.save(outDir + "outputCustomPaperSize.pdf");
		
		//Print the message
		System.out.println("ImplementCustomPaperSizeOfWorksheetForRendering executed successfully.");
	}
}
