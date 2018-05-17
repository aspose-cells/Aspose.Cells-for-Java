package AsposeCellsExamples.DrawingObjects;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class RotateTextWithShapeInsideWorksheet {
	
	static String srcDir = Utils.Get_SourceDirectory();
	static String outDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {

		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());

		//Load sample Excel file.
		Workbook wb = new Workbook(srcDir + "sampleRotateTextWithShapeInsideWorksheet.xlsx");
		 
		//Access first worksheet.
		Worksheet ws = wb.getWorksheets().get(0);
		 
		//Access cell B4 and add message inside it.
		Cell b4 = ws.getCells().get("B4");
		b4.putValue("Text is not rotating with shape because RotateTextWithShape is false.");
		 
		//Access first shape.
		Shape sh = ws.getShapes().get(0);
		 
		//Access shape text alignment.
		ShapeTextAlignment shapeTextAlignment = sh.getTextBody().getTextAlignment();
		 
		//Do not rotate text with shape by setting RotateTextWithShape as false.
		shapeTextAlignment.setRotateTextWithShape(false);
		 
		//Save the output Excel file.
		wb.save(outDir + "outputRotateTextWithShapeInsideWorksheet.xlsx");

		// Print the message
		System.out.println("RotateTextWithShapeInsideWorksheet executed successfully.");
	}
}
