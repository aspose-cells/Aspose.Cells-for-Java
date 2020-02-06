package AsposeCellsExamples.Worksheets;

import AsposeCellsExamples.Utils;
import com.aspose.cells.*;

public class GetCellValidationInODS {
	public static void main(String[] args) throws Exception {
		// ExStart: 1
		//directories
		String sourceDir = Utils.Get_SourceDirectory();

		// Instantiating a Workbook object
		Workbook workbook = new Workbook(sourceDir + "SampleBook1.ods");

		// Add a page break at cell Y30
		WorksheetCollection worksheets = workbook.getWorksheets();
		Worksheet worksheet = worksheets.get(0);

		Cell cell = worksheet.getCells().get("A9");

		if(cell.getValidation() != null)
		{
			System.out.println(cell.getValidation().getType());
		}
		// ExEnd:1

		System.out.println("GetCellValidationInODS executed successfully.");
	}
}
