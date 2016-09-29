package com.aspose.cells.examples.tables;

import com.aspose.cells.Cell;
import com.aspose.cells.Cells;
import com.aspose.cells.ListObject;
import com.aspose.cells.TableStyleType;
import com.aspose.cells.TotalsCalculation;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;
import com.aspose.cells.examples.RowsColumns.UnhidingRowsandColumns;

public class FormataListObject {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(FormataListObject.class) + "tables/";
		// Create a workbook.
		Workbook workbook = new Workbook();

		// Obtaining the reference of the default(first) worksheet
		Worksheet sheet = workbook.getWorksheets().get(0);

		// Obtaining Worksheet's cells collection
		Cells cells = sheet.getCells();

		// Setting the value to the cells
		Cell cell = cells.get("A1");
		cell.putValue("Employee");
		cell = cells.get("B1");
		cell.putValue("Quarter");
		cell = cells.get("C1");
		cell.putValue("Product");
		cell = cells.get("D1");
		cell.putValue("Continent");
		cell = cells.get("E1");
		cell.putValue("Country");
		cell = cells.get("F1");
		cell.putValue("Sale");

		cell = cells.get("A2");
		cell.putValue("David");
		cell = cells.get("A3");
		cell.putValue("David");
		cell = cells.get("A4");
		cell.putValue("David");
		cell = cells.get("A5");
		cell.putValue("David");
		cell = cells.get("A6");
		cell.putValue("James");
		cell = cells.get("A7");
		cell.putValue("James");
		cell = cells.get("A8");
		cell.putValue("James");
		cell = cells.get("A9");
		cell.putValue("James");
		cell = cells.get("A10");
		cell.putValue("James");
		cell = cells.get("A11");
		cell.putValue("Miya");
		cell = cells.get("A12");
		cell.putValue("Miya");
		cell = cells.get("A13");
		cell.putValue("Miya");
		cell = cells.get("A14");
		cell.putValue("Miya");
		cell = cells.get("A15");
		cell.putValue("Miya");

		cell = cells.get("B2");
		cell.putValue(1);
		cell = cells.get("B3");
		cell.putValue(2);
		cell = cells.get("B4");
		cell.putValue(3);
		cell = cells.get("B5");
		cell.putValue(4);
		cell = cells.get("B6");
		cell.putValue(1);
		cell = cells.get("B7");
		cell.putValue(2);
		cell = cells.get("B8");
		cell.putValue(3);
		cell = cells.get("B9");
		cell.putValue(4);
		cell = cells.get("B10");
		cell.putValue(4);
		cell = cells.get("B11");
		cell.putValue(1);
		cell = cells.get("B12");
		cell.putValue(1);
		cell = cells.get("B13");
		cell.putValue(2);
		cell = cells.get("B14");
		cell.putValue(2);
		cell = cells.get("B15");
		cell.putValue(2);

		cell = cells.get("C2");
		cell.putValue("Maxilaku");
		cell = cells.get("C3");
		cell.putValue("Maxilaku");
		cell = cells.get("C4");
		cell.putValue("Chai");
		cell = cells.get("C5");
		cell.putValue("Maxilaku");
		cell = cells.get("C6");
		cell.putValue("Chang");
		cell = cells.get("C7");
		cell.putValue("Chang");
		cell = cells.get("C8");
		cell.putValue("Chang");
		cell = cells.get("C9");
		cell.putValue("Chang");
		cell = cells.get("C10");
		cell.putValue("Chang");
		cell = cells.get("C11");
		cell.putValue("Geitost");
		cell = cells.get("C12");
		cell.putValue("Chai");
		cell = cells.get("C13");
		cell.putValue("Geitost");
		cell = cells.get("C14");
		cell.putValue("Geitost");
		cell = cells.get("C15");
		cell.putValue("Geitost");

		cell = cells.get("D2");
		cell.putValue("Asia");
		cell = cells.get("D3");
		cell.putValue("Asia");
		cell = cells.get("D4");
		cell.putValue("Asia");
		cell = cells.get("D5");
		cell.putValue("Asia");
		cell = cells.get("D6");
		cell.putValue("Europe");
		cell = cells.get("D7");
		cell.putValue("Europe");
		cell = cells.get("D8");
		cell.putValue("Europe");
		cell = cells.get("D9");
		cell.putValue("Europe");
		cell = cells.get("D10");
		cell.putValue("Europe");
		cell = cells.get("D11");
		cell.putValue("America");
		cell = cells.get("D12");
		cell.putValue("America");
		cell = cells.get("D13");
		cell.putValue("America");
		cell = cells.get("D14");
		cell.putValue("America");
		cell = cells.get("D15");
		cell.putValue("America");

		cell = cells.get("E2");
		cell.putValue("China");
		cell = cells.get("E3");
		cell.putValue("India");
		cell = cells.get("E4");
		cell.putValue("Korea");
		cell = cells.get("E5");
		cell.putValue("India");
		cell = cells.get("E6");
		cell.putValue("France");
		cell = cells.get("E7");
		cell.putValue("France");
		cell = cells.get("E8");
		cell.putValue("Germany");
		cell = cells.get("E9");
		cell.putValue("Italy");
		cell = cells.get("E10");
		cell.putValue("France");
		cell = cells.get("E11");
		cell.putValue("U.S.");
		cell = cells.get("E12");
		cell.putValue("U.S.");
		cell = cells.get("E13");
		cell.putValue("Brazil");
		cell = cells.get("E14");
		cell.putValue("U.S.");
		cell = cells.get("E15");
		cell.putValue("U.S.");

		cell = cells.get("F2");
		cell.putValue(2000);
		cell = cells.get("F3");
		cell.putValue(500);
		cell = cells.get("F4");
		cell.putValue(1200);
		cell = cells.get("F5");
		cell.putValue(1500);
		cell = cells.get("F6");
		cell.putValue(500);
		cell = cells.get("F7");
		cell.putValue(1500);
		cell = cells.get("F8");
		cell.putValue(800);
		cell = cells.get("F9");
		cell.putValue(900);
		cell = cells.get("F10");
		cell.putValue(500);
		cell = cells.get("F11");
		cell.putValue(1600);
		cell = cells.get("F12");
		cell.putValue(600);
		cell = cells.get("F13");
		cell.putValue(2000);
		cell = cells.get("F14");
		cell.putValue(500);
		cell = cells.get("F15");
		cell.putValue(900);

		// Adding a new List Object to the worksheet
		ListObject listObject = sheet.getListObjects().get(sheet.getListObjects().add("A1", "F15", true));

		// Adding Default Style to the table
		listObject.setTableStyleType(TableStyleType.TABLE_STYLE_MEDIUM_10);

		// Show Total
		listObject.setShowTotals(true);

		// Set the Quarter field's calculation type
		listObject.getListColumns().get(1).setTotalsCalculation(TotalsCalculation.COUNT);

		// Saving the Excel file
		workbook.save(dataDir + "FormataListObject_out.xlsx");
	}
}
