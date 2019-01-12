package AsposeCellsExamples.SmartMarkers;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class UsingFormulaParameterInSmartMarkerField {
	public static void main(String[] args) throws Exception {
		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());
		
		String outDir = Utils.Get_OutputDirectory();
		
		//Create array of strings that are actually Excel formulas
		String str1 =  "=\"01-This \" & \"is \" & \"concatenation\"";
		String str2 =  "=\"02-This \" & \"is \" & \"concatenation\"";
		String str3 =  "=\"03-This \" & \"is \" & \"concatenation\"";
		String str4 =  "=\"04-This \" & \"is \" & \"concatenation\"";
		String str5 =  "=\"05-This \" & \"is \" & \"concatenation\"";
		 
		String[] TestFormula = new String[]{str1, str2, str3, str4, str5};
		 
		//Create a workbook
		Workbook wb = new Workbook();
		 
		//Access first worksheet
		Worksheet ws = wb.getWorksheets().get(0);
		 
		//Put the smart marker field with formula parameter in cell A1
		Cells cells = ws.getCells();    
		Cell cell = cells.get("A1");
		cell.putValue("&=$Test(formula)");
		 
		//Create workbook designer, set data source and process it
		WorkbookDesigner wd = new WorkbookDesigner(wb);
		wd.setDataSource("Test", TestFormula);    
		wd.process();
		 
		//Save the workbook in xlsx format
		wb.save(outDir + "outputUsingFormulaParameterInSmartMarkerField.xlsx");

		//Print the message
		System.out.println("UsingFormulaParameterInSmartMarkerField executed successfully.");
	}
}
