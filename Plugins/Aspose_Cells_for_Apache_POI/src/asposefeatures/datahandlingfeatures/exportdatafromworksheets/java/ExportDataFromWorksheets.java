package asposefeatures.datahandlingfeatures.exportdatafromworksheets.java;

import java.io.FileInputStream;
import java.util.Arrays;

import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

public class ExportDataFromWorksheets
{
    public static void main(String[] args) throws Exception
    {
	String dataPath = "src/asposefeatures/datahandlingfeatures/exportdatafromworksheets/data/";

	// Creating a file stream containing the Excel file to be opened
	FileInputStream fstream = new FileInputStream(dataPath + "workbook.xls");

	// Instantiating a Workbook object
	Workbook workbook = new Workbook(fstream);

	// Accessing the first worksheet in the Excel file
	Worksheet worksheet = workbook.getWorksheets().get(0);

	// Exporting the contents of 7 rows and 2 columns starting from 1st cell
	// to Array.
	Object dataTable[][] = worksheet.getCells().exportArray(4, 0, 7, 8);

	for (int i = 0; i < dataTable.length; i++)
	{
	    System.out.println("[" + i + "]: " + Arrays.toString(dataTable[i]));
	}
	// Closing the file stream to free all resources
	fstream.close();
    }
}
