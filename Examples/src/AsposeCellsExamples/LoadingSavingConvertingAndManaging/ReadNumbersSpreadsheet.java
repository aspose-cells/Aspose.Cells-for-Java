package AsposeCellsExamples.LoadingSavingConvertingAndManaging;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class ReadNumbersSpreadsheet { 

	public static void main(String[] args) throws Exception {

		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());

		String srcDir = Utils.Get_SourceDirectory();
		String outDir = Utils.Get_OutputDirectory();

		// Specify load options, we want to load Numbers spreadsheet.
		LoadOptions opts = new LoadOptions(LoadFormat.NUMBERS);

		// Load the Numbers spreadsheet in workbook with above load options.
		Workbook wb = new Workbook(srcDir + "sampleNumbersByAppleInc.numbers", opts);

		// Save the workbook to pdf format
		wb.save(outDir + "outputNumbersByAppleInc.pdf", SaveFormat.PDF);

		// Print the message
		System.out.println("ReadNumbersSpreadsheet executed successfully.");
	}
}
