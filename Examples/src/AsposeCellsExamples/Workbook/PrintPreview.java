package AsposeCellsExamples.Workbook;

import AsposeCellsExamples.Utils;
import com.aspose.cells.*;

public class PrintPreview {

	public static void main(String[] args) throws Exception {

        // ExStart:1
        // The path to the directories.
        String sourceDir = Utils.Get_SourceDirectory();

        Workbook workbook = new Workbook(sourceDir + "Book1.xlsx");
        ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();
        WorkbookPrintingPreview preview = new WorkbookPrintingPreview(workbook, imgOptions);
        System.out.println("Workbook page count: " + preview.getEvaluatedPageCount());

        SheetPrintingPreview preview2 = new SheetPrintingPreview(workbook.getWorksheets().get(0), imgOptions );
        System.out.println("Worksheet page count: " + preview2.getEvaluatedPageCount());
        // ExEnd:1

		System.out.println("PrintPreview executed successfully.");
	}
}
