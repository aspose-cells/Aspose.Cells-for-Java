package featurescomparison.workingwithworkbook.createnewworkbook.java;

import com.aspose.cells.FileFormatType;
import com.aspose.cells.Workbook;

public class AsposeNewWorkbook
{
	public static void main(String[] args) throws Exception
	{
		String dataPath = "src/featurescomparison/workingwithworkbook/createnewworkbook/data/";
		
		Workbook workbook = new Workbook(); // Creating a Workbook object

		//Workbooks can be saved in many formats
		workbook.save(dataPath + "newWorkBook_Aspose_Out.xlsx", FileFormatType.XLSX);

		System.out.println("Worksheets are saved successfully."); // Print Message
	}
}