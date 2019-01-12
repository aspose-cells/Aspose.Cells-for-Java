package AsposeCellsExamples.Worksheets;

import com.aspose.cells.Workbook;

public class ClearAllPageBreaks {
	public static void main(String[] args) throws Exception {
		// Instantiating a Workbook object
		Workbook workbook = new Workbook();
		workbook.getWorksheets().get(0).getHorizontalPageBreaks().clear();
		workbook.getWorksheets().get(0).getVerticalPageBreaks().clear();
	}
}
