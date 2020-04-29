package AsposeCellsExamples.Data;

import AsposeCellsExamples.Utils;
import com.aspose.cells.*;

public class CreateUnionRange {
	public static void main(String[] args) throws Exception {
		// ExStart:1
		// Output directory
		String outputDir = Utils.Get_OutputDirectory();

		// Instantiating a Workbook object
		Workbook workbook = new Workbook();

		// Create union range
		UnionRange unionRange = workbook.getWorksheets().createUnionRange("sheet1!A1:A10,sheet1!C1:C10", 0);

		// Put value "ABCD" in the range
		unionRange.setValue("ABCD");

		// Saving the modified Excel file in default format
		workbook.save(outputDir + "CreateUnionRange_out.xlsx");
		// ExEnd:1

		System.out.println("CreateUnionRange executed successfully.");
	}
}
