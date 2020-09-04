package AsposeCellsExamples.WorkbookSettings;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class SupportNamedRangeFormulasInGermanLocale {
	
	static String srcDir = Utils.Get_SourceDirectory();
	static String outDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {
        // ExStart:1
		// Define variables
        String name = "HasFormula";
        String value = "=GET.CELL(48, INDIRECT(\"ZS\",FALSE))";

        // Load the template file
        Workbook wbSource = new Workbook(srcDir + "sampleNamedRangeTest.xlsm");
        
        // Get the worksheets collection
        WorksheetCollection wsCol = wbSource.getWorksheets();

        // Add new name to the names collection
        int nameIndex = wsCol.getNames().add(name);
        
        // Set value to the named range
        Name namedRange = wsCol.getNames().get(nameIndex);
        namedRange.setRefersTo(value);

        // Save the output file
        wbSource.save(outDir + "sampleOutputNamedRangeTest.xlsm");
        // ExEnd:1

		// Print the message
		System.out.println("SupportNamedRangeFormulasInGermanLocale executed successfully.");
	}
}