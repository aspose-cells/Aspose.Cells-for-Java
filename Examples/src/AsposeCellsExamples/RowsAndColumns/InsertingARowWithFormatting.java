package AsposeCellsExamples.RowsAndColumns;

import com.aspose.cells.CopyFormatType;
import com.aspose.cells.InsertOptions;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import AsposeCellsExamples.Utils;

public class InsertingARowWithFormatting {
	public static void main(String[] args) throws Exception {
		// ExStart:1
		String dataDir = Utils.getSharedDataDir(InsertingARowWithFormatting.class) + "RowsAndColumns/";
		// Instantiating a Workbook object
		Workbook workbook = new Workbook(dataDir + "Book1.xlsx");

		// Accessing the first worksheet in the Excel file
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Setting Formatting options
		InsertOptions insertOptions = new InsertOptions();
        insertOptions.setCopyFormatType(CopyFormatType.SAME_AS_ABOVE);
        
		// Inserting a row into the worksheet at 3rd position
        worksheet.getCells().insertRows(2, 1, insertOptions);

		// Saving the modified Excel file
		workbook.save(dataDir + "InsertingARowWithFormatting_out.xlsx");
		// ExEnd:1
		System.out.println("InsertingARowWithFormatting executed successfully!");
	}
}
