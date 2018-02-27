package AsposeCellsExamples.DrawingObjects;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class SetMarginsOfCommentOrShapeInsideTheWorksheet { 
	
	static String srcDir = Utils.Get_SourceDirectory();
	static String outDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {

		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());

		//Load the sample Excel file
		Workbook wb = new Workbook(srcDir + "sampleSetMarginsOfCommentOrShapeInsideTheWorksheet.xlsx");

		//Access first worksheet
		Worksheet ws = wb.getWorksheets().get(0);

		for(int idx =0; idx<ws.getShapes().getCount(); idx++)
		{
			//Access the shape
			Shape sh = ws.getShapes().get(idx);
			
			//Access the text alignment
			ShapeTextAlignment txtAlign = sh.getTextBody().getTextAlignment();

			//Set auto margin false
			txtAlign.setAutoMargin(false);

			//Set the top, left, bottom and right margins
			txtAlign.setTopMarginPt(10);
			txtAlign.setLeftMarginPt(10);
			txtAlign.setBottomMarginPt(10);
			txtAlign.setRightMarginPt(10);	    
		}

		//Save the output Excel file
		wb.save(outDir + "outputSetMarginsOfCommentOrShapeInsideTheWorksheet.xlsx");

		// Print the message
		System.out.println("SetMarginsOfCommentOrShapeInsideTheWorksheet executed successfully.");
	}
}
