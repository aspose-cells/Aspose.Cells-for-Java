package AsposeCellsExamples.Formulas;

import com.aspose.cells.Cell;
import com.aspose.cells.DateTime;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import AsposeCellsExamples.Utils;

public class RegisterAndCallFuncFromAddIn {
	static String srcDir = Utils.Get_SourceDirectory();
	static String outDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {
        // ExStart:
        // Create empty workbook
        Workbook workbook = new Workbook();

        // Register macro enabled add-in along with the function name
        int id = workbook.getWorksheets().registerAddInFunction(srcDir + "TESTUDF.xlam", "TEST_UDF", false);

        // Register more functions in the file (if any)
        workbook.getWorksheets().registerAddInFunction(id, "TEST_UDF1"); //in this way you can add more functions that are in the same file

        // Access first worksheet
        Worksheet worksheet = workbook.getWorksheets().get(0);

        // Access first cell
        Cell cell = worksheet.getCells().get("A1");

        // Set formula name present in the add-in
        cell.setFormula("=TEST_UDF()");

        // Save workbook to output XLSX format.
        workbook.save(outDir +  "test_udf.xlsx", SaveFormat.XLSX);
        // ExEnd:

		// Print message
		System.out.println("RegisterAndCallFuncFromAddIn executed successfully.");
	}
}
