package AsposeCellsExamples.Data;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class ApplyAdvancedFilterOfMicrosoftExcel {

	public static void main(String[] args) throws Exception {

		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());

		String srcDir = Utils.Get_SourceDirectory();
		String outDir = Utils.Get_OutputDirectory();

		// Load your source workbook
		Workbook wb = new Workbook(srcDir + "sampleAdvancedFilter.xlsx");

		// Access first worksheet
		Worksheet ws = wb.getWorksheets().get(0);

		// Apply advanced filter on range A5:D19 and criteria range is A1:D2
		// Besides, we want to filter in place
		// And, we want all filtered records not just unique records
		ws.advancedFilter(true, "A5:D19", "A1:D2", "", false);

		// Save the workbook in xlsx format
		wb.save(outDir + "outputAdvancedFilter.xlsx", SaveFormat.XLSX);

		// Print the message
		System.out.println("ApplyAdvancedFilterOfMicrosoftExcel executed successfully.");
	}
}
