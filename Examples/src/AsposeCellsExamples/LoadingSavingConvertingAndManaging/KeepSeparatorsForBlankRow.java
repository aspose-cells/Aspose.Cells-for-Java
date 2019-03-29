package AsposeCellsExamples.LoadingSavingConvertingAndManaging;

import com.aspose.cells.TxtSaveOptions;
import com.aspose.cells.Workbook;
import AsposeCellsExamples.Utils;
import AsposeCellsExamples.Tables.ConvertTableToRangeWithOptions;

public class KeepSeparatorsForBlankRow {

	public static void main(String[] args) throws Exception {
		// ExStart:1		
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ConvertTableToRangeWithOptions.class) + "LoadingSavingConvertingAndManaging/";
		// Open an existing file that contains a table/list object in it
		Workbook workbook = new Workbook(dataDir + "KeepSeparatorsForBlankRow.xlsx");

		// Instantiate Text File's Save Options
        TxtSaveOptions options = new TxtSaveOptions();
        
        // Set KeepSeparatorsForBlankRow to true show separators in blank rows
        options.setKeepSeparatorsForBlankRow(true);
        
        // Save the file with the options
        workbook.save(dataDir + "KeepSeparatorsForBlankRow.out.csv", options);
		// ExEnd:1

        System.out.println("KeepSeparatorsForBlankRow executed successfully.");
	}
}
