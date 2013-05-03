/* 
 * Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Cells. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
 
package programmersguide.asposecells.creatingcharts.fundamentalfeatures.changechartpositionandsize.java;

import com.aspose.cells.*;

public class ChangeChartPositionAndSize
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = "src/programmersguide/asposecells/creatingcharts/changechartpositionandsize/data/";
        
        String filePath = dataDir + "book1.xls";

        Workbook workbook = new Workbook(filePath);

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
        workbook.save(dataDir + "book1.out.xls");
        
        // Print message
        System.out.println("Position and Size of Chart is changed successfully.\nPlease check data folder for output.");
    }
}
