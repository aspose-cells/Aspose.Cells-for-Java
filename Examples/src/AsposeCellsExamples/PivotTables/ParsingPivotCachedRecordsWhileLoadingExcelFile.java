package AsposeCellsExamples.PivotTables;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class ParsingPivotCachedRecordsWhileLoadingExcelFile {
	
	static String srcDir = Utils.Get_SourceDirectory();
	static String outDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {

		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());
		
		//Create load options
		LoadOptions options = new LoadOptions();

		//Set ParsingPivotCachedRecords true, default value is false
		options.setParsingPivotCachedRecords(true); 

		//Load the sample Excel file containing pivot table cached records
		Workbook wb = new Workbook(srcDir + "sampleParsingPivotCachedRecordsWhileLoadingExcelFile.xlsx", options);

		//Access first worksheet
		Worksheet ws = wb.getWorksheets().get(0);

		//Access first pivot table
		PivotTable pt = ws.getPivotTables().get(0);

		//Set refresh data flag true
		pt.setRefreshDataFlag(true);

		//Refresh and calculate pivot table
		pt.refreshData();
		pt.calculateData();

		//Set refresh data flag false
		pt.setRefreshDataFlag(false);

		//Save the output Excel file
		wb.save(outDir + "outputParsingPivotCachedRecordsWhileLoadingExcelFile.xlsx");
		
		// Print the message
		System.out.println("ParsingPivotCachedRecordsWhileLoadingExcelFile executed successfully.");
	}
}
