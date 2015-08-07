package featurescomparison.workingwithworksheets.converttocsv.java;

import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;

public class AsposeConvertWorkbookToCSV 
{
	public static void main(String[] args) throws Exception 
	{
		String dataPath = "src/featurescomparison/workingwithworksheets/converttocsv/data/";
		
		//Instantiate a new workbook with Excel file path
		Workbook workbook = new Workbook(dataPath + "workbook.xls");

		//Save the document in PDF format
		workbook.save(dataPath + "AsposeWorkbookCSV.csv", SaveFormat.CSV);
		
		System.out.println("Done.");
	}
}
