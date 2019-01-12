package AsposeCellsExamples.TechnicalArticles;

import com.aspose.cells.Workbook;

public class SetAutoRecoverProperty {
	public static void main(String[] args) throws Exception {
		// Create workbook object
		Workbook workbook = new Workbook();

		// Read AutoRecover property
		System.out.println("AutoRecover: " + workbook.getSettings().getAutoRecover());

		// Set AutoRecover property to false
		workbook.getSettings().setAutoRecover(false);

		// Save the workbook
		workbook.save("SetAutoRecoverProperty_out.xlsx");

		// Read the saved workbook again
		workbook = new Workbook("SetAutoRecoverProperty_out.xlsx");

		// Read AutoRecover property
		System.out.println("AutoRecover: " + workbook.getSettings().getAutoRecover());

	}
}
