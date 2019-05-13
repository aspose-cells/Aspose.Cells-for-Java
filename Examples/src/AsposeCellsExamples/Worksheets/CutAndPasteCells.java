package AsposeCellsExamples.Worksheets;

import com.aspose.cells.Range;
import com.aspose.cells.ShiftType;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

import AsposeCellsExamples.Utils;
public class CutAndPasteCells {
	
	public static void main(String[] args) throws Exception {
		// ExStart:1
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(CountNumberOfCells.class) + "Worksheets/";
		// Instantiating a Workbook object
		Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.getWorksheets().get(0);

        worksheet.getCells().get(0, 2).setValue(1);
        worksheet.getCells().get(1, 2).setValue(2);
        worksheet.getCells().get(2, 2).setValue(3);
        worksheet.getCells().get(2, 3).setValue(4);
        worksheet.getCells().createRange(0, 2, 3, 1).setName("NamedRange");

        Range cut = worksheet.getCells().createRange("C:C");
        worksheet.getCells().insertCutCells(cut, 0, 1, ShiftType.RIGHT);
        workbook.save(dataDir + "CutAndPasteCells.xlsx");
        // ExEnd:1

		System.out.println("CutAndPasteCells executed successfully.");
	}
}
