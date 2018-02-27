package AsposeCellsExamples.Data;

import java.util.ArrayList;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class SpecifyFormulaFieldsWhileImportingDataToWorksheet {
	
	static String outDir = Utils.Get_OutputDirectory();
	
	//User-defined class to hold data items
	public class DataItems
	{
		private int m_Number1;
		private int m_Number2;
		private String m_Formula1;
		private String m_Formula2;
		
		public DataItems(int num1, int num2, String form1, String form2)
		{
			this.m_Number1 = num1;
			this.m_Number2 = num2;
			this.m_Formula1 = form1;
			this.m_Formula2 = form2;
		}

		public int getNumber1()
		{
			return this.m_Number1;
		}
		
		public int getNumber2()
		{
			return this.m_Number2;
		}
		
		public String getFormula1()
		{
			return this.m_Formula1;
		}
		
		public String getFormula2()
		{
			return this.m_Formula2;
		}
		
	}//DataItems


	public void Run() throws Exception
	{
		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());
		
		//List to hold data items
		ArrayList<DataItems> dis = new ArrayList<DataItems>();
		
		//Define 1st data item and add it in list
		int num1 = 2002;
		int num2 = 3502;
		String form1 = "=SUM(A2,B2)";
		String form2 = "=HYPERLINK(\"https://www.aspose.com\",\"Aspose Website\")";	
		DataItems di = new DataItems(num1, num2, form1, form2);
		dis.add(di);
		
		//Define 2nd data item and add it in list
		num1 = 2003;
		num2 = 3503;
		form1 = "=SUM(A3,B3)";
		form2 = "=HYPERLINK(\"https://www.aspose.com\",\"Aspose Website\")";	
		di = new DataItems(num1, num2, form1, form2);
		dis.add(di);

		//Define 3rd data item and add it in list
		num1 = 2004;
		num2 = 3504;
		form1 = "=SUM(A4,B4)";
		form2 = "=HYPERLINK(\"https://www.aspose.com\",\"Aspose Website\")";	
		di = new DataItems(num1, num2, form1, form2);
		dis.add(di);

		//Define 4th data item and add it in list
		num1 = 2005;
		num2 = 3505;
		form1 = "=SUM(A5,B5)";
		form2 = "=HYPERLINK(\"https://www.aspose.com\",\"Aspose Website\")";	
		di = new DataItems(num1, num2, form1, form2);
		dis.add(di);
		
		//Create workbook object
		Workbook wb = new Workbook();

		//Access first worksheet
		Worksheet ws = wb.getWorksheets().get(0);

		//Specify import table options
		ImportTableOptions opts = new ImportTableOptions();

		//Specify which field is formula field, here the last two fields are formula fields
		//opts.setColumnIndexes(new int[] {3, 0, 2, 1});
		opts.setFormulas(new boolean[] {false, false, true, true });

		//Import custom objects
		ws.getCells().importCustomObjects(dis, 0, 0, opts);

		//Calculate formula
		wb.calculateFormula();

		//Autofit columns
		ws.autoFitColumns();

		//Save the output Excel file
		wb.save(outDir + "outputSpecifyFormulaFieldsWhileImportingDataToWorksheet.xlsx");
		
		// Print the message
		System.out.println("SpecifyFormulaFieldsWhileImportingDataToWorksheet executed successfully.");
	}

	public static void main(String[] args) throws Exception {
		new SpecifyFormulaFieldsWhileImportingDataToWorksheet().Run();
	}
}
