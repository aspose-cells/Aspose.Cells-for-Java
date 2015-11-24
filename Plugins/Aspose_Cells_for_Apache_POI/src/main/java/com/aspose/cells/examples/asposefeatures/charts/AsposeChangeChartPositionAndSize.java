package com.aspose.cells.examples.asposefeatures.charts;

import com.aspose.cells.Chart;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class AsposeChangeChartPositionAndSize
{
    public static void main(String[] args) throws Exception
    {
	// The path to the documents directory.
        String dataDir = Utils.getDataDir(AsposeChangeChartPositionAndSize.class);

	Workbook workbook = new Workbook(dataDir + "AsposeChart.xls");

	Worksheet worksheet = workbook.getWorksheets().get(0);

	//Load the chart from source worksheet
	Chart chart = worksheet.getCharts().get(0);

	//Resize the chart
	chart.getChartObject().setWidth(400);
	chart.getChartObject().setHeight(300);

	//Reposition the chart
	chart.getChartObject().setX(250);
	chart.getChartObject().setY(150);

	//Output the file
	workbook.save(dataDir + "AsposeChangeChart.xlsx");
	
	System.out.println("Chart Size changed and repositioned.");
    }
}
