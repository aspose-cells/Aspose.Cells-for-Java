package featurescomparison.workingwithworksheets.zoomfactor.java;

import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.WorksheetCollection;

public class AsposeZoom
{
	public static void main(String[] args) throws Exception
	{
		String dataPath = "src/featurescomparison/workingwithworksheets/zoomfactor/data/";
		
		//Instantiating a Excel object by excel file path
		Workbook workbook = new Workbook();

		//Accessing the first worksheet in the Excel file
		WorksheetCollection worksheets = workbook.getWorksheets();
		Worksheet worksheet = worksheets.get(0);

		//Setting the zoom factor of the worksheet to 75     
		worksheet.setZoom(75);

		//Saving the modified Excel file in default format
		workbook.save(dataPath + "AsposeZoom_Out.xls");
		
		System.out.println("Aspose Zoom Created.");
	}
}
