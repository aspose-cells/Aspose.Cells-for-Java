package AsposeCellsExamples.Worksheets;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class UtilizeSheet_SheetId_PropertyOfOpenXml {
	
	static String srcDir = Utils.Get_SourceDirectory();
	static String outDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {

		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());
		
		//Load source Excel file
		Workbook wb = new Workbook(srcDir + "sampleSheetId.xlsx");
		  
		//Access first worksheet
		Worksheet ws = wb.getWorksheets().get(0);
		  
		//Print its Sheet or Tab Id on console
		System.out.println("Sheet or Tab Id: " + ws.getTabId());
		  
		//Change Sheet or Tab Id
		ws.setTabId(358);
		  
		//Save the workbook
		wb.save(outDir + "outputSheetId.xlsx");
		
		// Print the message
		System.out.println("UtilizeSheet_SheetId_PropertyOfOpenXml executed successfully.");
	}
}
