package AsposeCellsExamples.WorkbookSettings;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class SupportNamedRangeFormulasInGermanLocale {
	
	static String srcDir = Utils.Get_SourceDirectory();
	static String outDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {

        // ExStart:1
        String name = "HasFormula";
        String value = "=GET.CELL(48, INDIRECT(\"ZS\",FALSE))";

        Workbook wbSource = new Workbook(srcDir + "sampleNamedRangeTest.xlsm");
        WorksheetCollection wsCol = wbSource.getWorksheets();

        int nameIndex = wsCol.getNames().add(name);
        Name namedRange = wsCol.getNames().get(nameIndex);
        namedRange.setRefersTo(value);

        wbSource.save(outDir + "sampleOutputNamedRangeTest.xlsm");
        // ExEnd:1

		// Print the message
		System.out.println("SupportNamedRangeFormulasInGermanLocale executed successfully.");
	}
}
