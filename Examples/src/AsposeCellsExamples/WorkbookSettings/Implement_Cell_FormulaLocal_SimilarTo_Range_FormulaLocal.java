package AsposeCellsExamples.WorkbookSettings;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class Implement_Cell_FormulaLocal_SimilarTo_Range_FormulaLocal {
	
	static String srcDir = Utils.Get_SourceDirectory();
	static String outDir = Utils.Get_OutputDirectory();
	
	//Implement GlobalizationSettings class
	class GS extends GlobalizationSettings {
	 
		public String getLocalFunctionName(String standardName)
		{
			//Change the SUM function name as per your needs.
			if(standardName.equals("SUM"))
			{
				return "UserFormulaLocal_SUM";				
			}
					
			//Change the AVERAGE function name as per your needs.
			if (standardName.equals("AVERAGE"))
			{
				return "UserFormulaLocal_AVERAGE";
			}
					
			return "";
		}//getLocalFunctionName
	}//GS extends GlobalizationSettings

	public void Run() throws Exception {

		//Create workbook
		Workbook wb = new Workbook();

		//Assign GlobalizationSettings implementation class
		wb.getSettings().setGlobalizationSettings(new GS());

		//Access first worksheet
		Worksheet ws = wb.getWorksheets().get(0);

		//Access some cell
		Cell cell = ws.getCells().get("C4");

		//Assign SUM formula and print its FormulaLocal
		cell.setFormula("SUM(A1:A2)");
		System.out.println("Formula Local: " + cell.getFormulaLocal());

		//Assign AVERAGE formula and print its FormulaLocal
		cell.setFormula("=AVERAGE(B1:B2, B5)");
		System.out.println("Formula Local: " + cell.getFormulaLocal());
	}

	public static void main(String[] args) throws Exception {

		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());
		
		Implement_Cell_FormulaLocal_SimilarTo_Range_FormulaLocal pg = new Implement_Cell_FormulaLocal_SimilarTo_Range_FormulaLocal();
		pg.Run();

		// Print the message
		System.out.println("Implement_Cell_FormulaLocal_SimilarTo_Range_FormulaLocal executed successfully.");
	}
}
