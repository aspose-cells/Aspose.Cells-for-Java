package AsposeCellsExamples.DrawingObjects;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class ExtractTextFromGearTypeSmartArtShape { 
	
	static String srcDir = Utils.Get_SourceDirectory();
	static String outDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {

		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());
		
		// Load sample Excel file containing gear type smart art shape.
		Workbook wb = new Workbook(srcDir + "sampleExtractTextFromGearTypeSmartArtShape.xlsx");

		// Access first worksheet.
		Worksheet ws = wb.getWorksheets().get(0);

		// Access first shape.
		Shape sh = ws.getShapes().get(0);

		// Get the result of gear type smart art shape in the form of group shape.
		GroupShape gs = sh.getResultOfSmartArt();

		// Get the list of individual shapes consisting of group shape.
		Shape[] shps = gs.getGroupedShapes();

		// Extract the text of gear type shapes and print them on console.
		for (int i = 0; i < shps.length; i++)
		{
			Shape s = shps[i];

			if (s.getType() == AutoShapeType.GEAR_9 || s.getType() == AutoShapeType.GEAR_6)
			{
				System.out.println("Gear Type Shape Text: " + s.getText());
			}
		}//for
		
		// Print the message
		System.out.println("ExtractTextFromGearTypeSmartArtShape executed successfully.");
	}
}
