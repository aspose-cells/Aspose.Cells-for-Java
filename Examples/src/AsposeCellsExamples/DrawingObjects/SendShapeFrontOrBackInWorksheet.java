package AsposeCellsExamples.DrawingObjects;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class SendShapeFrontOrBackInWorksheet {
	
	static String srcDir = Utils.Get_SourceDirectory();
	static String outDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {

		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());
		
		//Load source Excel file
		Workbook wb = new Workbook(srcDir + "sampleToFrontOrBack.xlsx");
		 
		//Access first worksheet
		Worksheet ws = wb.getWorksheets().get(0);
		 
		//Access first and fourth shape
		Shape sh1 = ws.getShapes().get(0);
		Shape sh4 = ws.getShapes().get(3);
		 
		//Print the Z-Order position of the shape
		System.out.println("Z-Order Shape 1: " + sh1.getZOrderPosition());
		 
		//Send this shape to front
		sh1.toFrontOrBack(2);
		 
		//Print the Z-Order position of the shape
		System.out.println("Z-Order Shape 4: " + sh4.getZOrderPosition());
		 
		//Send this shape to back
		sh4.toFrontOrBack(-2);
		 
		//Save the output Excel file
		wb.save(outDir + "outputToFrontOrBack.xlsx");

		// Print the message
		System.out.println("SendShapeFrontOrBackInWorksheet executed successfully.");
	}
}
