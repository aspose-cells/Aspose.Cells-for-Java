package com.aspose.cells.examples.articles;

import com.aspose.cells.Chart;
import com.aspose.cells.ChartType;
import com.aspose.cells.DataLabels;
import com.aspose.cells.DataLablesSeparatorType;
import com.aspose.cells.FileFormatType;
import com.aspose.cells.ImageFormat;
import com.aspose.cells.ImageOrPrintOptions;
import com.aspose.cells.LabelPositionType;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class CreatePieChartWithLeaderLines {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(CreatePieChartWithLeaderLines.class) + "articles/";
		// Create an instance of Workbook in XLSX format
		Workbook workbook = new Workbook(FileFormatType.XLSX);

		// Access the first worksheet
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Add two columns of data
		worksheet.getCells().get("A1").putValue("Retail");
		worksheet.getCells().get("A2").putValue("Services");
		worksheet.getCells().get("A3").putValue("Info & Communication");
		worksheet.getCells().get("A4").putValue("Transport Equip");
		worksheet.getCells().get("A5").putValue("Construction");
		worksheet.getCells().get("A6").putValue("Other Products");
		worksheet.getCells().get("A7").putValue("Wholesale");
		worksheet.getCells().get("A8").putValue("Land Transport");
		worksheet.getCells().get("A9").putValue("Air Transport");
		worksheet.getCells().get("A10").putValue("Electric Appliances");
		worksheet.getCells().get("A11").putValue("Securities");
		worksheet.getCells().get("A12").putValue("Textiles & Apparel");
		worksheet.getCells().get("A13").putValue("Machinery");
		worksheet.getCells().get("A14").putValue("Metal Products");
		worksheet.getCells().get("A15").putValue("Cash");
		worksheet.getCells().get("A16").putValue("Banks");

		worksheet.getCells().get("B1").putValue(10.4);
		worksheet.getCells().get("B2").putValue(5.2);
		worksheet.getCells().get("B3").putValue(6.4);
		worksheet.getCells().get("B4").putValue(10.4);
		worksheet.getCells().get("B5").putValue(7.9);
		worksheet.getCells().get("B6").putValue(4.1);
		worksheet.getCells().get("B7").putValue(3.5);
		worksheet.getCells().get("B8").putValue(5.7);
		worksheet.getCells().get("B9").putValue(3);
		worksheet.getCells().get("B10").putValue(14.7);
		worksheet.getCells().get("B11").putValue(3.6);
		worksheet.getCells().get("B12").putValue(2.8);
		worksheet.getCells().get("B13").putValue(7.8);
		worksheet.getCells().get("B14").putValue(2.4);
		worksheet.getCells().get("B15").putValue(1.8);
		worksheet.getCells().get("B16").putValue(10.1);

		// Create a pie chart and add it to the collection of charts
		int id = worksheet.getCharts().add(ChartType.PIE, 3, 3, 23, 13);
		// Access newly created Chart instance
		Chart chart = worksheet.getCharts().get(id);

		// Set series data range
		chart.getNSeries().add("B1:B16", true);
		// Set category data range
		chart.getNSeries().setCategoryData("A1:A16");
		// Turn off legend
		chart.setShowLegend(false);

		// Access data labels
		DataLabels dataLabels = chart.getNSeries().get(0).getDataLabels();
		// Turn on category names
		dataLabels.setShowCategoryName(true);
		// Turn on percentage format
		dataLabels.setShowPercentage(true);
		// Set position
		dataLabels.setPosition(LabelPositionType.OUTSIDE_END);
		// Set separator
		dataLabels.setSeparator(DataLablesSeparatorType.COMMA);
		
		//Turn on leader lines
		chart.getNSeries().get(0).setHasLeaderLines(true);

		//Calculate chart
		chart.calculate();

		//You need to move DataLabels a little leftward or rightward depending on their position
		//to show leader lines
		int DELTA = 100;
		for (int i = 0; i < chart.getNSeries().get(0).getPoints().getCount(); i++)
		{
		    int X = chart.getNSeries().get(0).getPoints().get(i).getDataLabels().getX();
		    //If it is greater than 2000, then move the X position a little right
		    //otherwise move the X position a little left
		    if (X > 2000)
		        chart.getNSeries().get(0).getPoints().get(i).getDataLabels().setX(X + DELTA);
		    else
		    	chart.getNSeries().get(0).getPoints().get(i).getDataLabels().setX(X - DELTA);
		}
		
		//In order to save the chart image, create an instance of ImageOrPrintOptions
		ImageOrPrintOptions anOption = new ImageOrPrintOptions();
		//Set image format
		anOption.setImageFormat(ImageFormat.getPng());
		//Set resolution
		anOption.setHorizontalResolution(200);
		anOption.setVerticalResolution(200);

		//Render chart to image
		chart.toImage(dataDir + "CPieChartWLLines_out.png", anOption);

		//Save the workbook to see chart inside the Excel
		workbook.save(dataDir + "CPieChartWLLines_out.xlsx");
		

	}
}
