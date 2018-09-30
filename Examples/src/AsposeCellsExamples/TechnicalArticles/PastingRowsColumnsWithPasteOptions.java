package AsposeCellsExamples.TechnicalArticles;

import com.aspose.cells.CopyOptions;
import com.aspose.cells.PasteOptions;
import com.aspose.cells.PasteType;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import AsposeCellsExamples.Utils;

public class PastingRowsColumnsWithPasteOptions {
	// The path to the documents directory.
	static String srcDir = Utils.Get_SourceDirectory();
	static String outDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {

		// ExStart:ContentCopyForAccessibilityOption
		// Load some excel file
		Workbook wb = new Workbook(srcDir + "book1.xlsx");

		// Access the first sheet which contains chart
		Worksheet source = wb.getWorksheets().get(0);

		// Add another sheet named DestSheet
		Worksheet destination = wb.getWorksheets().add("DestSheet");

		// Set CopyOptions.ReferToDestinationSheet to true
		CopyOptions options = new CopyOptions();
		options.setReferToDestinationSheet(true);

		// Set PasteOptions
		PasteOptions pasteOptions = new PasteOptions();
		pasteOptions.setPasteType(PasteType.VALUES);
		pasteOptions.setOnlyVisibleCells(true);

		// Copy all the rows of source worksheet to destination worksheet which includes chart as well
		// The chart data source will now refer to DestSheet
		destination.getCells().copyRows(source.getCells(), 0, 0, source.getCells().getMaxDisplayRange().getRowCount(), options, pasteOptions);

		// Save workbook in xlsx format
		wb.save(outDir + "destination.xlsx", SaveFormat.XLSX);
		
		// Print message
		System.out.println("Pasting Rows Columns With Paste Options performed successfully.");
		
		// ExStart:ContentCopyForAccessibilityOption
	}
}
