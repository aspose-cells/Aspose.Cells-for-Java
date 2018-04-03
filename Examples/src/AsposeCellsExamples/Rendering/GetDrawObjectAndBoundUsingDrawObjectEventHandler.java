package AsposeCellsExamples.Rendering;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class GetDrawObjectAndBoundUsingDrawObjectEventHandler { 
	
	static String srcDir = Utils.Get_SourceDirectory();
	static String outDir = Utils.Get_OutputDirectory();
	
	//Implement the concrete class of DrawObjectEventHandler
	class clsDrawObjectEventHandler extends DrawObjectEventHandler
	{
		public void draw(DrawObject drawObject, float x, float y, float width, float height)
		{
			System.out.println();

			//Print the coordinates and the value of Cell object
			if (drawObject.getType() == DrawObjectEnum.CELL)
			{
				System.out.println("[X]: " + x + " [Y]: " + y + " [Width]: " + width + " [Height]: " + height + " [Cell Value]: " + drawObject.getCell().getStringValue());
			}

			//Print the coordinates and the shape name of Image object
			if (drawObject.getType() == DrawObjectEnum.IMAGE)
			{
				System.out.println("[X]: " + x + " [Y]: " + y + " [Width]: " + width + " [Height]: " + height + " [Shape Name]: " + drawObject.getShape().getName());
			}

			System.out.println("----------------------");
		}
	}

	void Run() throws Exception
	{
		//Load sample Excel file
		Workbook wb = new Workbook(srcDir + "sampleGetDrawObjectAndBoundUsingDrawObjectEventHandler.xlsx");
	 
		//Specify Pdf save options
		PdfSaveOptions opts = new PdfSaveOptions();
	 
		//Assign the instance of DrawObjectEventHandler class
		opts.setDrawObjectEventHandler(new clsDrawObjectEventHandler());
	 
		//Save to Pdf format with Pdf save options
		wb.save(outDir + "outputGetDrawObjectAndBoundUsingDrawObjectEventHandler.pdf", opts);
	}
	
	public static void main(String[] args) throws Exception {

		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());
		
		new GetDrawObjectAndBoundUsingDrawObjectEventHandler().Run();

		// Print the message
		System.out.println("GetDrawObjectAndBoundUsingDrawObjectEventHandler executed successfully.");
	}
}
