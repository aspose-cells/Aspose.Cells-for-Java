package AsposeCellsExamples.LoadingSavingConvertingAndManaging;

import AsposeCellsExamples.Utils;
import com.aspose.cells.*;

public class ConvertActiveWorksheetToSVG {

	public static void main(String[] args) throws Exception {
		// ExStart: 1
		String outputDir = Utils.Get_OutputDirectory();

		// Instantiate a workbook
		Workbook workbook = new Workbook();

		// Put sample text in the first cell of first worksheet in the newly created workbook
		workbook.getWorksheets().get(0).getCells().get("A1").setValue("DEMO TEXT ON SHEET1");

		// Add second worksheet in the workbook
		workbook.getWorksheets().add(SheetType.WORKSHEET);

		// Set text in first cell of the second sheet
		workbook.getWorksheets().get(1).getCells().get("A1").setValue("DEMO TEXT ON SHEET2");

		// Set currently active sheet index to 1 i.e. Sheet2
		workbook.getWorksheets().setActiveSheetIndex(1);

		// Save workbook to SVG. It shall render the active sheet only to SVG
		workbook.save(outputDir + "ConvertActiveWorksheetToSVG_out.svg");
		// ExEnd: 1

		// Print message
		System.out.println("ConvertActiveWorksheetToSVG executed successfully.");

	}
}
