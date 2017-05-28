package AsposeCellsExamples.LoadingSavingConvertingAndManaging;

import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;
import AsposeCellsExamples.Utils;

public class QuickExcelToXPSConversion {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ExportWholeWorkbookToXPS.class) + "LoadingSavingConvertingAndManaging/";
		// Open an Excel file
		Workbook workbook = new Workbook(dataDir + "Book1.xlsx");

		// Save in XPS format
		workbook.save("QEToXPSConversion_out.xps", SaveFormat.XPS);
	}
}
