package AsposeCellsExamples.WorkbookSettings;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class FindMaximumRowsAndColumnsSupportedByXLSAndXLSXFormats { 
	
	static String srcDir = Utils.Get_SourceDirectory();
	static String outDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {

		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());
		
		// Print message about XLS format.
		System.out.println("Maximum Rows and Columns supported by XLS format.");

		// Create workbook in XLS format.
		Workbook wb = new Workbook(FileFormatType.EXCEL_97_TO_2003);

		// Print the maximum rows and columns supported by XLS format.
		int maxRows = wb.getSettings().getMaxRow() + 1;
		int maxCols = wb.getSettings().getMaxColumn() + 1;
		System.out.println("Maximum Rows: " + maxRows);
		System.out.println("Maximum Columns: " + maxCols);
		System.out.println();

		// Print message about XLSX format.
		System.out.println("Maximum Rows and Columns supported by XLSX format.");

		// Create workbook in XLSX format.
		wb = new Workbook(FileFormatType.XLSX);

		// Print the maximum rows and columns supported by XLSX format.
		maxRows = wb.getSettings().getMaxRow() + 1;
		maxCols = wb.getSettings().getMaxColumn() + 1;
		System.out.println("Maximum Rows: " + maxRows);
		System.out.println("Maximum Columns: " + maxCols);

		// Print the message
		System.out.println("FindMaximumRowsAndColumnsSupportedByXLSAndXLSXFormats executed successfully.");
	}
}
