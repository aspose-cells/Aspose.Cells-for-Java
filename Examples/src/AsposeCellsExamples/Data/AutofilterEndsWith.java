package AsposeCellsExamples.Data;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class AutofilterEndsWith {

	// The path to the documents directory.
	static String srcDir = Utils.Get_SourceDirectory();
	static String outDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {

        // ExStart:1
        // Instantiating a Workbook object containing sample data
        Workbook workbook = new Workbook(srcDir + "sourseSampleCountryNames.xlsx");

        // Accessing the first worksheet in the Excel file
        Worksheet worksheet = workbook.getWorksheets().get(0);

        // Creating AutoFilter by giving the cells range
        worksheet.getAutoFilter().setRange("A1:A18");

        // Initialize filter for rows starting with string "Ba"
        worksheet.getAutoFilter().custom(0, FilterOperatorType.ENDS_WITH, "ia");

        //Refresh the filter to show/hide filtered rows
        worksheet.getAutoFilter().refresh();

        // Saving the modified Excel file
        workbook.save(outDir +  "outSourseSampleCountryNames.xlsx");
        // ExEnd:1
        
		// Print message
		System.out.println("AutofilterEndsWith executed successfully");
	}
}
