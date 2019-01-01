package AsposeCellsExamples.Data;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class CheckIfValidationInCellDropDown {

	static String srcDir = Utils.Get_SourceDirectory();

	public static void main(String[] args) throws Exception {

        Workbook book = new Workbook(srcDir + "sampleValidation.xlsx");
        Worksheet sheet = book.getWorksheets().get("Sheet1");
        Cells cells  = sheet.getCells();
        Cell a2 = cells.get("A2");
        Validation va2 = a2.getValidation();
        if(va2.getInCellDropDown()) {
            System.out.println("A2 is a dropdown");
        } else {
            System.out.println("A2 is NOT a dropdown");
        }
        Cell b2 = cells.get("B2");
        Validation vb2 = b2.getValidation();
        if(vb2.getInCellDropDown()) {
            System.out.println("B2 is a dropdown");
        } else {
            System.out.println("B2 is NOT a dropdown");
        }
        Cell c2 = cells.get("C2");
        Validation vc2 = c2.getValidation();
        if(vc2.getInCellDropDown()) {
            System.out.println("C2 is a dropdown");
        } else {
            System.out.println("C2 is NOT a dropdown");
        }

		// Print message
		System.out.println("CheckIfValidationInCellDropDown completed successfully");

	}
}
