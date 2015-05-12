/* 
 * Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Cells. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
 
package programmersguide.asposecells.creatingcharts.fundamentalfeatures.createchart.java;

import com.aspose.cells.*;

public class CreateChart
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = "src/programmersguide/asposecells/creatingcharts/fundamentalfeatures/createchart/data/";
        
      //Instantiating a Workbook object
        Workbook workbook = new Workbook();

        //Obtaining the reference of the first worksheet
        WorksheetCollection worksheets = workbook.getWorksheets();
        Worksheet sheet =  worksheets.get(0);

        //Adding some sample value to cells
        Cells cells = sheet.getCells();
        Cell cell = cells.get("A1");
        cell.setValue(50);
        cell = cells.get("A2");
        cell. setValue (100);
        cell = cells.get("A3");
        cell.setValue(150);
        cell = cells.get("B1");
        cell.setValue(4);
        cell = cells.get("B2");
        cell.setValue(20);
        cell = cells.get("B3");
        cell.setValue(50);

        ChartCollection charts = sheet.getCharts();

        //Adding a chart to the worksheet
        int chartIndex = charts.add(ChartType.PYRAMID,5,0,15,5);
        Chart chart = charts.get(chartIndex);

        //Adding NSeries (chart data source) to the chart ranging from "A1" cell to "B3"
        SeriesCollection serieses = chart.getNSeries();
        serieses.add("A1:B3", true);

        //Saving the Excel file
        workbook.save(dataDir + "book1.out.xls");
        
        // Print message
        System.out.println("Workbook with chart is successfully created.");
        
    }
}