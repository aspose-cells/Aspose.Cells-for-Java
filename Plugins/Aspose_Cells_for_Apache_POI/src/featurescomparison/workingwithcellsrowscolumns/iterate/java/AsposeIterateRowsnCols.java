package featurescomparison.workingwithcellsrowscolumns.iterate.java;

import com.aspose.cells.Cell;
import com.aspose.cells.CellValueType;
import com.aspose.cells.Cells;
import com.aspose.cells.ColumnCollection;
import com.aspose.cells.Range;
import com.aspose.cells.Row;
import com.aspose.cells.RowCollection;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;

public class AsposeIterateRowsnCols
{
	public static void main(String[] args) throws Exception
	{
		String dataPath = "src/featurescomparison/workingwithcellsrowscolumns/iterate/data/";
		
		Workbook workbook = new Workbook(dataPath + "workbook.xls");
		
		//Accessing the first worksheet in the Excel file
		Worksheet worksheet = workbook.getWorksheets().get(0);
		Cells cells = worksheet.getCells();
		
		//Access the Maximum Display Range
		Range range = worksheet.getCells().getMaxDisplayRange();
		int tcols = range.getColumnCount();
		int trows = range.getRowCount();
		
		System.out.println("Total Rows:" + trows);
		System.out.println("Total Cols:" + tcols);

		RowCollection rows = cells.getRows();
				
		for (int i = 0 ; i < rows.getCount() ; i++)
		{
			for (int j = 0 ; j < tcols ; j++)
			{
				System.out.print(cells.get(i,j).getName() + " - " + cells.get(i,j).getValue() + "\t");
			}
			System.out.println("");
		}
	}
}