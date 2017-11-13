package AsposeCellsExamples.Formulas;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class InterruptOrCancelFormulaCalculationOfWorkbook { 
	
	static String srcDir = Utils.Get_SourceDirectory();
	static String outDir = Utils.Get_OutputDirectory();

	//Implement calculation monitor class
	class clsCalculationMonitor extends AbstractCalculationMonitor
	{
	    public void beforeCalculate(int sheetIndex, int rowIndex, int colIndex)
	    {
	        //Find the cell name
	        String cellName = CellsHelper.cellIndexToName(rowIndex, colIndex);
	  
	        //Print the sheet, row and column index as well as cell name
	        System.out.println(sheetIndex + "----" + rowIndex + "----" + colIndex + "----" + cellName);
	  
	        //If cell name is B8, interrupt/cancel the formula calculation
	        if (cellName.equals("B8") == true)
	        {
	            this.interrupt("Interrupt/Cancel the formula calculation");
	        }//if
	  
	    }//beforeCalculate
	  
	}//clsCalculationMonitor
	  
	//---------------------------------------------------------     
	//---------------------------------------------------------
	  
	public void Run() throws Exception
	{   
	    //Load the sample Excel file
	    Workbook wb = new Workbook(srcDir + "sampleCalculationMonitor.xlsx");
	 
	    //Create calculation options and assign instance of calculation monitor class
	    CalculationOptions opts = new CalculationOptions();
	    opts.setCalculationMonitor(new clsCalculationMonitor());
	 
	    //Calculate formula with calculation options
	    wb.calculateFormula(opts);
	}
	
	public static void main(String[] args) throws Exception {

		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());
		
		InterruptOrCancelFormulaCalculationOfWorkbook pg = new InterruptOrCancelFormulaCalculationOfWorkbook();
		pg.Run();
			
		// Print the message
		System.out.println("InterruptOrCancelFormulaCalculationOfWorkbook executed successfully.");
	}
}
