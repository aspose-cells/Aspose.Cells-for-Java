package com.aspose.cells.examples.articles;

import com.aspose.cells.Cell;
import com.aspose.cells.Cells;
import com.aspose.cells.Chart;
import com.aspose.cells.ChartType;
import com.aspose.cells.Color;
import com.aspose.cells.ComboBox;
import com.aspose.cells.MsoDrawingType;
import com.aspose.cells.Range;
import com.aspose.cells.Style;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class UsingDynamicFormulas {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(UsingDynamicFormulas.class) + "articles/";

		// Create a workbook object
		Workbook workbook = new Workbook();

		// Get the first worksheet
		Worksheet sheet = workbook.getWorksheets().get(0);

		// Access cells collection of first worksheet
		Cells cells = sheet.getCells();

		// Create a range in the second worksheet
		Range range = cells.createRange("C21", "C24");

		// Name the range
		range.setName("MyRange");

		// Fill different cells with data in the range
		range.get(0, 0).putValue("North");
		range.get(1, 0).putValue("South");
		range.get(2, 0).putValue("East");
		range.get(3, 0).putValue("West");

		ComboBox comboBox = (ComboBox) sheet.getShapes().addShape(MsoDrawingType.COMBO_BOX, 15, 0, 2, 0, 17, 64);
		comboBox.setInputRange("=MyRange");
		comboBox.setLinkedCell("=B16");
		comboBox.setSelectedIndex(0);
		Cell cell = cells.get("B16");
		Style style = cell.getStyle();
		style.getFont().setColor(Color.getWhite());
		cell.setStyle(style);

		cells.get("C16").setFormula("=INDEX(Sheet1!$C$21:$C$24,$B$16,1)");

		// Put some data for chart source
		// Data Headers
		cells.get("D15").putValue("Jan");
		cells.get("D20").putValue("Jan");

		cells.get("E15").putValue("Feb");
		cells.get("E20").putValue("Feb");

		cells.get("F15").putValue("Mar");
		cells.get("F20").putValue("Mar");

		cells.get("G15").putValue("Apr");
		cells.get("G20").putValue("Apr");

		cells.get("H15").putValue("May");
		cells.get("H20").putValue("May");

		cells.get("I15").putValue("Jun");
		cells.get("I20").putValue("Jun");

		// Data
		cells.get("D21").putValue(304);
		cells.get("D22").putValue(402);
		cells.get("D23").putValue(321);
		cells.get("D24").putValue(123);

		cells.get("E21").putValue(300);
		cells.get("E22").putValue(500);
		cells.get("E23").putValue(219);
		cells.get("E24").putValue(422);

		cells.get("F21").putValue(222);
		cells.get("F22").putValue(331);
		cells.get("F23").putValue(112);
		cells.get("F24").putValue(350);

		cells.get("G21").putValue(100);
		cells.get("G22").putValue(200);
		cells.get("G23").putValue(300);
		cells.get("G24").putValue(400);

		cells.get("H21").putValue(200);
		cells.get("H22").putValue(300);
		cells.get("H23").putValue(400);
		cells.get("H24").putValue(500);

		cells.get("I21").putValue(400);
		cells.get("I22").putValue(200);
		cells.get("I23").putValue(200);
		cells.get("I24").putValue(100);

		// Dynamically load data on selection of Dropdown value
		cells.get("D16").setFormula("=IFERROR(VLOOKUP($C$16,$C$21:$I$24,2,FALSE),0)");
		cells.get("E16").setFormula("=IFERROR(VLOOKUP($C$16,$C$21:$I$24,3,FALSE),0)");
		cells.get("F16").setFormula("=IFERROR(VLOOKUP($C$16,$C$21:$I$24,4,FALSE),0)");
		cells.get("G16").setFormula("=IFERROR(VLOOKUP($C$16,$C$21:$I$24,5,FALSE),0)");
		cells.get("H16").setFormula("=IFERROR(VLOOKUP($C$16,$C$21:$I$24,6,FALSE),0)");
		cells.get("I16").setFormula("=IFERROR(VLOOKUP($C$16,$C$21:$I$24,7,FALSE),0)");

		// Create Chart
		int index = sheet.getCharts().add(ChartType.COLUMN, 0, 3, 12, 9);
		Chart chart = sheet.getCharts().get(index);
		chart.getNSeries().add("='Sheet1'!$D$16:$I$16", false);
		chart.getNSeries().get(0).setName("=C16");
		chart.getNSeries().setCategoryData("=$D$15:$I$15");

		// Save result on disc
		workbook.save(dataDir + "UDynamicFormulas_out.xlsx");

	}
}
