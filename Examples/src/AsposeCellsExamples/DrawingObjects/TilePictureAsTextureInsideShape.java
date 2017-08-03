package AsposeCellsExamples.DrawingObjects;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class TilePictureAsTextureInsideShape {

	public static void main(String[] args) throws Exception {

		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());

		String srcDir = Utils.Get_SourceDirectory();
		String outDir = Utils.Get_OutputDirectory();

		// Load sample Excel file
		Workbook wb = new Workbook(srcDir + "sampleTextureFill_IsTiling.xlsx");

		// Access first worksheet
		Worksheet ws = wb.getWorksheets().get(0);

		// Access first shape inside the worksheet
		Shape sh = ws.getShapes().get(0);

		// Tile Picture as a Texture inside the Shape
		sh.getFill().getTextureFill().setTiling(true);

		// Save the output Excel file
		wb.save(outDir + "outputTextureFill_IsTiling.xlsx");

		// Print the message
		System.out.println("TilePictureAsTextureInsideShape executed successfully.");
	}
}
