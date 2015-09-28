package featurescomparison.workingwithworkbook.openandsaveworkbook.java;

import com.aspose.cells.Workbook;

/**
 * @author Shoaib Khan
 */

public class OpenSaveSpreadSheetAspose
{
	public static void main(String[] args) throws Exception
	{
		String dataPath = "src/featurescomparison/workingwithworkbook/openandsaveworkbook/data/";
		
		// Creating an Workbook object with an Excel file path
		Workbook workbook = new Workbook(dataPath + "pivot.xlsm");

		// Saving the Excel file
		workbook.save(dataPath + "pivot-rtt-Aspose.xlsm");

		// Print Message
		System.out.println("Worksheet saved successfully.");
	}
}
