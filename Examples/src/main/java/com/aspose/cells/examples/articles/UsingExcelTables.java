package com.aspose.cells.examples.articles;

import com.aspose.cells.Cells;
import com.aspose.cells.Chart;
import com.aspose.cells.ChartType;
import com.aspose.cells.ListObjectCollection;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class UsingExcelTables {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(UsingExcelTables.class) + "articles/";
		// Create an instance of Workbook
		Workbook book = new Workbook();
		// Access first worksheet from the collection
		Worksheet sheet = book.getWorksheets().get(0);
		// Access cells collection of the first worksheet
		Cells cells = sheet.getCells();

		// Insert data column wise
		cells.get("A1").putValue("Category");
		cells.get("A2").putValue("Fruit");
		cells.get("A3").putValue("Fruit");
		cells.get("A4").putValue("Fruit");
		cells.get("A5").putValue("Fruit");
		cells.get("A6").putValue("Vegetables");
		cells.get("A7").putValue("Vegetables");
		cells.get("A8").putValue("Vegetables");
		cells.get("A9").putValue("Vegetables");
		cells.get("A10").putValue("Beverages");
		cells.get("A11").putValue("Beverages");
		cells.get("A12").putValue("Beverages");

		cells.get("B1").putValue("Food");
		cells.get("B2").putValue("Apple");
		cells.get("B3").putValue("Banana");
		cells.get("B4").putValue("Apricot");
		cells.get("B5").putValue("Grapes");
		cells.get("B6").putValue("Carrot");
		cells.get("B7").putValue("Onion");
		cells.get("B8").putValue("Cabage");
		cells.get("B9").putValue("Potatoe");
		cells.get("B10").putValue("Coke");
		cells.get("B11").putValue("Coladas");
		cells.get("B12").putValue("Fizz");

		cells.get("C1").putValue("Cost");
		cells.get("C2").putValue(2.2);
		cells.get("C3").putValue(3.1);
		cells.get("C4").putValue(4.1);
		cells.get("C5").putValue(5.1);
		cells.get("C6").putValue(4.4);
		cells.get("C7").putValue(5.4);
		cells.get("C8").putValue(6.5);
		cells.get("C9").putValue(5.3);
		cells.get("C10").putValue(3.2);
		cells.get("C11").putValue(3.6);
		cells.get("C12").putValue(5.2);

		cells.get("D1").putValue("Profit");
		cells.get("D2").putValue(0.1);
		cells.get("D3").putValue(0.4);
		cells.get("D4").putValue(0.5);
		cells.get("D5").putValue(0.6);
		cells.get("D6").putValue(0.7);
		cells.get("D7").putValue(1.3);
		cells.get("D8").putValue(0.8);
		cells.get("D9").putValue(1.3);
		cells.get("D10").putValue(2.2);
		cells.get("D11").putValue(2.4);
		cells.get("D12").putValue(3.3);

		// Create ListObject. Get the List objects collection in the first worksheet
		ListObjectCollection listObjects = sheet.getListObjects();

		// Add a List based on the data source range with headers on
		int index = listObjects.add(0, 0, 11, 3, true);

		sheet.autoFitColumns();

		// Create chart based on ListObject
		index = sheet.getCharts().add(ChartType.COLUMN, 21, 1, 35, 18);
		Chart chart = sheet.getCharts().get(index);
		chart.setChartDataRange("A1:D12", true);
		chart.getNSeries().setCategoryData("A2:B12");

		// Calculate chart
		chart.calculate();

		// Save spreadsheet
		book.save(dataDir + "UsingExcelTables_out.xlsx");

	}
}
