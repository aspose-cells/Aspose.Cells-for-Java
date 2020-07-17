package AsposeCellsExamples.Worksheets;

import AsposeCellsExamples.Utils;
import com.aspose.cells.SheetType;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

public class DetectInternationalMacroSheet {

	public static void main(String[] args) throws Exception {

		// ExStart:1
		// The path to the Source directory.
        String sourceDir = Utils.Get_SourceDirectory();

		//Load source Excel file
        Workbook workbook = new Workbook(sourceDir + "InternationalMacroSheet.xlsm");

        //Get Sheet Type
        int sheetType = workbook.getWorksheets().get(0).getType();

        //Print Sheet Type
        if(sheetType == SheetType.INTERNATIONAL_MACRO) {
            System.out.println("Sheet Type: INTERNATIONAL_MACRO");
        } else {
            System.out.println("Sheet Type: Other");
        }
        // ExEnd:1

        System.out.println("DetectInternationalMacroSheet executed successfully.");
	}
}
