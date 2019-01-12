package AsposeCellsExamples.LoadingSavingConvertingAndManaging;

import com.aspose.cells.TxtSaveOptions;
import com.aspose.cells.Workbook;
import AsposeCellsExamples.Utils;

public class SavingTextFilewithCustomSeparator {

	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(SavingTextFilewithCustomSeparator.class) + "LoadingSavingConvertingAndManaging/";

		// Creating an Workbook object with an Excel file path
		Workbook workbook = new Workbook(dataDir + "Book1.xlsx");

		TxtSaveOptions toptions = new TxtSaveOptions();
		// Specify the separator
		toptions.setSeparator(';');
		workbook.save(dataDir + "STFWCSeparator_out.csv");

		// Print Message
		System.out.println("Worksheets are saved successfully.");

	}
}
