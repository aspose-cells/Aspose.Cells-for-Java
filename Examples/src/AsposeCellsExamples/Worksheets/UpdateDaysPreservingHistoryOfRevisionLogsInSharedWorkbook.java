package AsposeCellsExamples.Worksheets;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class UpdateDaysPreservingHistoryOfRevisionLogsInSharedWorkbook {
	
	static String srcDir = Utils.Get_SourceDirectory();
	static String outDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {

		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());
		
		//Create empty workbook
		Workbook wb = new Workbook();

		//Share the workbook
		wb.getSettings().setShared(true);

		//Update DaysPreservingHistory of RevisionLogs
		wb.getWorksheets().getRevisionLogs().setDaysPreservingHistory(7);

		//Save the workbook
		wb.save(outDir + "outputShared_DaysPreservingHistory.xlsx");

		// Print the message
		System.out.println("UpdateDaysPreservingHistoryOfRevisionLogsInSharedWorkbook executed successfully.");
	}
}
