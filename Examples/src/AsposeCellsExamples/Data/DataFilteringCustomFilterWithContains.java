package AsposeCellsExamples.Data;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class DataFilteringCustomFilterWithContains {

	// The path to the documents directory.
	static String srcDir = Utils.Get_SourceDirectory();
	static String outDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {

		// ExStart:DataFilteringCustomFilterWithContains
        // Instantiating a Workbook object containing sample data
        Workbook workbook = new Workbook(srcDir + "sourseSampleCountryNames.xlsx");

        // Accessing the first worksheet in the Excel file
        Worksheet worksheet = workbook.getWorksheets().get(0);

        // Creating AutoFilter by giving the cells range
        worksheet.getAutoFilter().setRange("A1:A18");

        // Initialize filter for rows containing string "Ba"
        worksheet.getAutoFilter().custom(0, FilterOperatorType.CONTAINS, "Ba");

        //Refresh the filter to show/hide filtered rows
        worksheet.getAutoFilter().refresh();

        // Saving the modified Excel file
        workbook.save(outDir + "outSourseSampleCountryNames.xlsx");

		// Print message
		System.out.println("Data Filtering custom filter with contains completed successfully");
		// ExEnd:DataFilteringCustomFilterWithContains
	}
}
