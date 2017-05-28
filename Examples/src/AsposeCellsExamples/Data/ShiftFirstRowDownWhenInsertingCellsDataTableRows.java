package AsposeCellsExamples.Data;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class ShiftFirstRowDownWhenInsertingCellsDataTableRows {
	
	class CellsDataTable implements ICellsDataTable
	{
	    //This is the current row index
	    int m_index=-1;
	     
	    //These are your column names
	    String[] colsNames = new String[] { "Pet", "Fruit", "Country", "Color" };
	 
	    //These are the data of each column
	    String[] col0data = new String[] { "Dog", "Cat", "Duck" };
	    String[] col1data = new String[] { "Apple", "Pear", "Banana" };
	    String[] col2data = new String[] { "UK", "USA", "China" };
	    String[] col3data = new String[] { "Red", "Green", "Blue" };
	 
	    //Combine all of the data into a single two dimensional array
	    String[][] colsData = new String[][]{ col0data, col1data, col2data, col3data};
	 
	 
	    public void beforeFirst() {
	        m_index = -1;
	    }
	 
	    public Object get(int columnIndex) {
	         
	        Object o = null;
	        o = colsData[columnIndex][m_index];
	        return o;
	    }
	 
	    public Object get(String columnName) {
	        return null;
	    }
	 
	    public String[] getColumns() {
	        return colsNames;
	    }
	 
	    public int getCount() {
	        return col0data.length;
	    }
	 
	    public boolean next() {
	        m_index++;
	        return true;
	    }
	     
	}//End Class - CellsDataTable
	
	public void Run() throws Exception
	{
		String srcDir = Utils.Get_SourceDirectory();
		String outDir = Utils.Get_OutputDirectory();
		
	    //Create the instance of Cells Data Table
	    CellsDataTable cellsDataTable = new CellsDataTable();
	 
	    //Load the sample workbook
	    Workbook wb = new Workbook(srcDir + "sampleImportTableOptionsShiftFirstRowDown.xlsx");
	 
	    //Access first worksheet
	    Worksheet ws = wb.getWorksheets().get(0);
	 
	    //Import data table options
	    ImportTableOptions opts = new ImportTableOptions();
	 
	    //We do now want to shift the first row down when inserting rows. 
	    opts.setShiftFirstRowDown(false);
	 
	    //Import cells data table 
	    ws.getCells().importData(cellsDataTable, 2, 2, opts);
	 
	    //Save the workbook
	    wb.save(outDir + "outputImportTableOptionsShiftFirstRowDown-False.xlsx");
	}
	
	public static void main(String[] args) throws Exception {
		
		ShiftFirstRowDownWhenInsertingCellsDataTableRows pg = new ShiftFirstRowDownWhenInsertingCellsDataTableRows();
		pg.Run();
		
		//Print the message
		System.out.println("ShiftFirstRowDownWhenInsertingCellsDataTableRows executed successfully.");
	}
}
