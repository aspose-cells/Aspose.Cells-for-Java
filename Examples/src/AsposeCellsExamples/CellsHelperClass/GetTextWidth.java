package AsposeCellsExamples.CellsHelperClass;

import AsposeCellsExamples.Utils;
import com.aspose.cells.Cell;
import com.aspose.cells.CellsHelper;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

public class GetTextWidth {

	public static void main(String[] args) throws Exception {
		// ExStart:1
		// The path to the output directory.
		String sourceDir = Utils.Get_SourceDirectory();Workbook workbook = new Workbook(sourceDir + "GetTextWidthSample.xlsx");

		System.out.println("Text width: " + CellsHelper.getTextWidth(workbook.getWorksheets().get(0).getCells().get("A1").getStringValue(), workbook.getDefaultStyle().getFont(), 1));
		// ExEnd:1

		System.out.println("GetTextWidth executed successfully.");
	}
}
