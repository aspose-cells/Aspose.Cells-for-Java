package AsposeCellsExamples.Data;

import AsposeCellsExamples.Utils;
import com.aspose.cells.CellArea;
import com.aspose.cells.Validation;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

public class AddValidationArea {

	public static void main(String[] args) throws Exception {

		// ExStart:1
		// The path to the directories.
		String sourceDir = Utils.Get_SourceDirectory();
		String outputDir = Utils.Get_OutputDirectory();

		Workbook workbook = new Workbook(sourceDir + "ValidationsSample.xlsx");

		// Access first worksheet.
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Accessing the Validations collection of the worksheet
		Validation validation = worksheet.getValidations().get(0);

		// Create your cell area.
		CellArea cellArea = CellArea.createCellArea("D5", "E7");

		// Adding the cell area to Validation
		validation.addArea(cellArea, false, false);

		// Save the output workbook.
		workbook.save(outputDir + "ValidationsSample_out.xlsx");
		// ExEnd:1

		System.out.println("AddValidationArea executed successfully.");

	}
}
