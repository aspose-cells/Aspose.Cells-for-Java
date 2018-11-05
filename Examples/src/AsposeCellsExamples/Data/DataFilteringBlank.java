package AsposeCellsExamples.Data;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class DataFilteringBlank {

	// The path to the documents directory.
	static String srcDir = Utils.Get_SourceDirectory();
	static String outDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {

		// ExStart:1
        // Instantiating a Workbook object
        // Opening the Excel file through the file stream
        Workbook workbook = new Workbook(srcDir + "Blank.xlsx");

        // Accessing the first worksheet in the Excel file
        Worksheet worksheet = workbook.getWorksheets().get(0);

        // Call matchBlanks function to apply the filter
        worksheet.getAutoFilter().matchBlanks(0);

        // Call refresh function to update the worksheet
        worksheet.getAutoFilter().refresh();

        // Saving the modified Excel file
        workbook.save(outDir + "FilteredBlank.xlsx");
		// ExEnd:1
        
		// Print message
		System.out.println("Data Filtering Blank Process completed successfully");
	}
}
