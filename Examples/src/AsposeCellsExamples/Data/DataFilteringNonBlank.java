package AsposeCellsExamples.Data;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class DataFilteringNonBlank {

	// The path to the documents directory.
	static String srcDir = Utils.Get_SourceDirectory();
	static String outDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {

		// ExStart:DataFilteringNonBlank
        // Instantiating a Workbook object
        // Opening the Excel file through the file stream
        Workbook workbook = new Workbook(srcDir + "NonBlank.xlsx");

        // Accessing the first worksheet in the Excel file
        Worksheet worksheet = workbook.getWorksheets().get(0);

        // Call matchBlanks function to apply the filter
        worksheet.getAutoFilter().matchBlanks(0);

        // Call refresh function to update the worksheet
        worksheet.getAutoFilter().refresh();

        // Saving the modified Excel file
        workbook.save(outDir + "FilteredNonBlank.xlsx");

		// Print message
		System.out.println("Data Filtering Non Blank Process completed successfully");
		// ExEnd:DataFilteringNonBlank
	}
}
