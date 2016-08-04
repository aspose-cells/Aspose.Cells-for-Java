package com.aspose.cells.examples.files.utility;

import com.aspose.cells.Cell;
import com.aspose.cells.Cells;
import com.aspose.cells.PivotTable;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Style;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class IsPivotTableCompatibleForExcel2003 {

	public static void main(String[] args) throws Exception {
		
		// The path to the resource directory.
		String dataDir = Utils.getSharedDataDir(ExpandTextFromRightToLeftWhileExportingExcelFileToHTML.class) + "Conversion/";
				
		//Load source excel file containing sample pivot table
		Workbook wb = new Workbook(dataDir + "sample-pivot-table.xlsx");

		//Access first worksheet that contains pivot table data
		Worksheet dataSheet = wb.getWorksheets().get(0);

		//Access cell A3 and sets its data
		Cells cells = dataSheet.getCells();
		Cell cell = cells.get("A3");
		cell.putValue("FooBar");

		//Access cell B3 and sets its data
		//We set B3 a very long string which has more than 255 characters
		String longStr = "Very long text 1. very long text 2. very long text 3. very long text 4. very long text 5. very long text 6. very long text 7. very long text 8. very long text 9. very long text 10. very long text 11. very long text 12. very long text 13. very long text 14. very long text 15. very long text 16. very long text 17. very long text 18. very long text 19. very long text 20. End of text.";
		cell = cells.get("B3");
		cell.putValue(longStr);

		//Print the length of cell B3 string
		System.out.println("Length of original data string: " + cell.getStringValue().length());

		//Access cell C3 and sets its data
		cell = cells.get("C3");
		cell.putValue("closed");

		//Access cell D3 and sets its data
		cell = cells.get("D3");
		cell.putValue("2016/07/21");

		//Access the second worksheet that contains pivot table
		Worksheet pivotSheet = wb.getWorksheets().get(1);

		//Access the pivot table
		PivotTable pivotTable = pivotSheet.getPivotTables().get(0);

		//IsExcel2003Compatible property tells if PivotTable is compatible for Excel2003 while refreshing PivotTable.
		//If it is true, a string must be less than or equal to 255 characters, so if the string is greater than 255 characters,
		//it will be truncated. If false, a string will not have the aforementioned restriction.
		//The default value is true.
		pivotTable.setExcel2003Compatible(true);
		pivotTable.refreshData();
		pivotTable.calculateData();

		//Check the value of cell B5 of pivot sheet.
		//It will be 255 because we have set IsExcel2003Compatible property to true
		//All the data after 255 characters has been truncated
		Cell b5 = pivotSheet.getCells().get("B5");
		System.out.println("Length of cell B5 after setting IsExcel2003Compatible property to True: " + b5.getStringValue().length());

		//Now set IsExcel2003Compatible property to false and again refresh
		pivotTable.setExcel2003Compatible(false);
		pivotTable.refreshData();
		pivotTable.calculateData();

		//Now it will print 383 the original length of cell data
		//The data has not been truncated now.
		b5 = pivotSheet.getCells().get("B5");
		System.out.println("Length of cell B5 after setting IsExcel2003Compatible property to False: " + b5.getStringValue().length());

		//Set the row height and column width of cell B5 and also wrap its text
		pivotSheet.getCells().setRowHeight(b5.getRow(), 100);
		pivotSheet.getCells().setColumnWidth(b5.getColumn(), 65);
		Style st = b5.getStyle();
		st.setTextWrapped(true);
		b5.setStyle(st);

		//Save workbook in xlsx format
		wb.save(dataDir + "output.xlsx", SaveFormat.XLSX);

	}

}
