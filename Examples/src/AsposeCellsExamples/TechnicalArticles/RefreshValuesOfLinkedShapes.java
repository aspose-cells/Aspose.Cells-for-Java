package AsposeCellsExamples.TechnicalArticles;

import com.aspose.cells.Cell;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import AsposeCellsExamples.Utils;

public class RefreshValuesOfLinkedShapes {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(RefreshValuesOfLinkedShapes.class) + "TechnicalArticles/";
		// Create workbook from source file
		Workbook workbook = new Workbook(dataDir + "LinkedShape.xlsx");

		// Access first worksheet
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Change the value of cell A1
		Cell cell = worksheet.getCells().get("A1");
		cell.putValue(100);

		// Update the value of the Linked Picture which is linked to cell A1
		worksheet.getShapes().updateSelectedValue();

		// Save the workbook in pdf format
		workbook.save(dataDir + "RVOfLinkedShapes_out.pdf", SaveFormat.PDF);

	}
}
