package com.aspose.cells.examples.articles;

import java.io.FileInputStream;

import com.aspose.cells.Cells;
import com.aspose.cells.Chart;
import com.aspose.cells.ChartType;
import com.aspose.cells.Color;
import com.aspose.cells.Legend;
import com.aspose.cells.LegendPositionType;
import com.aspose.cells.SheetType;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;
import java.io.*;

public class SetPictureAsBackgroundFillInChart {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(SetPictureAsBackgroundFillInChart.class) + "articles/";
		// Create a new Workbook.
		Workbook workbook = new Workbook();

		// Get the first worksheet.
		Worksheet sheet = workbook.getWorksheets().get(0);

		// Set the name of worksheet
		sheet.setName("Data");

		// Get the cells collection in the sheet.
		Cells cells = workbook.getWorksheets().get(0).getCells();

		// Put some values into a cells of the Data sheet.
		cells.get("A1").putValue("Region");
		cells.get("A2").putValue("France");
		cells.get("A3").putValue("Germany");
		cells.get("A4").putValue("England");
		cells.get("A5").putValue("Sweden");
		cells.get("A6").putValue("Italy");
		cells.get("A7").putValue("Spain");
		cells.get("A8").putValue("Portugal");
		cells.get("B1").putValue("Sale");
		cells.get("B2").putValue(70000);
		cells.get("B3").putValue(55000);
		cells.get("B4").putValue(30000);
		cells.get("B5").putValue(40000);
		cells.get("B6").putValue(35000);
		cells.get("B7").putValue(32000);
		cells.get("B8").putValue(10000);

		// Add a chart sheet.
		int sheetIndex = workbook.getWorksheets().add(SheetType.CHART);
		sheet = workbook.getWorksheets().get(sheetIndex);

		// Set the name of worksheet
		sheet.setName("Chart");

		// Create chart
		int chartIndex = 0;
		chartIndex = sheet.getCharts().add(ChartType.COLUMN, 1, 1, 25, 10);
		Chart chart = sheet.getCharts().get(chartIndex);

		// Set some properties of chart plot area. To set a picture as fill format and make the border invisible.

		File file = new File(dataDir + "aspose-logo.png");
		byte[] data = new byte[(int) file.length()];
		FileInputStream fis = new FileInputStream(file);
		fis.read(data);

		chart.getPlotArea().getArea().getFillFormat().setImageData(data);
		chart.getPlotArea().getBorder().setVisible(false);

		// Set properties of chart title
		chart.getTitle().setText("Sales By Region");
		chart.getTitle().getFont().setColor(Color.getBlue());
		chart.getTitle().getFont().setBold(true);
		chart.getTitle().getFont().setSize(12);

		// Set properties of nseries
		chart.getNSeries().add("Data!B2:B8", true);
		chart.getNSeries().setCategoryData("Data!A2:A8");
		chart.getNSeries().setColorVaried(true);

		// Set the Legend.
		Legend legend = chart.getLegend();
		legend.setPosition(LegendPositionType.TOP);

		// Save the excel file
		workbook.save(dataDir + "SPAsBFillInChart_out.xls");

	}
}
