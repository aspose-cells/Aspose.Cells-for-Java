package com.aspose.cells.examples.asposefeatures.charts;

import java.io.FileOutputStream;

import com.aspose.cells.Cells;
import com.aspose.cells.Chart;
import com.aspose.cells.ChartPoint;
import com.aspose.cells.ChartPointCollection;
import com.aspose.cells.ChartType;
import com.aspose.cells.Color;
import com.aspose.cells.ImageFormat;
import com.aspose.cells.ImageOrPrintOptions;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class AsposeChartToImage
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(AsposeChartToImage.class);

        //Create a new Workbook.
        Workbook workbook = new Workbook();

        //Get the first worksheet.
        Worksheet sheet = workbook.getWorksheets().get(0);

        //Set the name of worksheet
        sheet.setName("Data");

        //Get the cells collection in the sheet.
        Cells cells = workbook.getWorksheets().get(0).getCells();

        //Put some values into a cells of the Data sheet.
        cells.get("A1").setValue("Region");
        cells.get("A2").setValue("France");
        cells.get("A3").setValue("Germany");
        cells.get("A4").setValue("England");
        cells.get("A5").setValue("Sweden");
        cells.get("A6").setValue("Italy");
        cells.get("A7").setValue("Spain");
        cells.get("A8").setValue("Portugal");
        cells.get("B1").setValue("Sale");
        cells.get("B2").setValue(70000);
        cells.get("B3").setValue(55000);
        cells.get("B4").setValue(30000);
        cells.get("B5").setValue(40000);
        cells.get("B6").setValue(35000);
        cells.get("B7").setValue(32000);
        cells.get("B8").setValue(10000);

        //Create chart
        int chartIndex = sheet.getCharts().add(ChartType.COLUMN, 12, 1, 33, 12);
        Chart chart = sheet.getCharts().get(chartIndex);

        //Set properties of chart title
        chart.getTitle().setText("Sales By Region");
        chart.getTitle().getTextFont().setBold(true);
        chart.getTitle().getTextFont().setSize(12);

        //Set properties of nseries
        chart.getNSeries().add("Data!B2:B8", true);
        chart.getNSeries().setCategoryData("Data!A2:A8");

        //Set the fill colors for the series's data points (France - Portugal(7 points))
        ChartPointCollection chartPoints = chart.getNSeries().get(0).getPoints();

        ChartPoint point = chartPoints.get(0);
        point.getArea().setForegroundColor(Color.getCyan());

        point = chartPoints.get(1);
        point.getArea().setForegroundColor(Color.getBlue());

        point = chartPoints.get(2);
        point.getArea().setForegroundColor(Color.getYellow());

        point = chartPoints.get(3);
        point.getArea().setForegroundColor(Color.getRed());

        point = chartPoints.get(4);
        point.getArea().setForegroundColor(Color.getBlack());

        point = chartPoints.get(5);
        point.getArea().setForegroundColor(Color.getGreen());

        point = chartPoints.get(6);
        point.getArea().setForegroundColor(Color.getMaroon());

        //Set the legend invisible
        chart.setShowLegend(false);

        //Get the Chart image
        ImageOrPrintOptions imgOpts = new ImageOrPrintOptions();
        imgOpts.setImageFormat(ImageFormat.getPng());

        //Save the chart image file.
        chart.toImage(new FileOutputStream(dataDir + "AsposeChartImage.png"), imgOpts);
    }	
}
