package AsposeCellsExamples.DrawingObjects;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class DetermineIfShapeIsSmartArtShape { 
	
	static String srcDir = Utils.Get_SourceDirectory();
	static String outDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {

		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());
				
		//Load the sample smart art shape - Excel file
		Workbook wb = new Workbook(srcDir + "sampleSmartArtShape.xlsx");
		  
		//Access first worksheet
		Worksheet ws = wb.getWorksheets().get(0);
		  
		//Access first shape
		Shape sh = ws.getShapes().get(0);
		  
		//Determine if shape is smart art
		System.out.println("Is Smart Art Shape: " + sh.isSmartArt());

		// Print the message
		System.out.println("DetermineIfShapeIsSmartArtShape executed successfully.");
	}
}
