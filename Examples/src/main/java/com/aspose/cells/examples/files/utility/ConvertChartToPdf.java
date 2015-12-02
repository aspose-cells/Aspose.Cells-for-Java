/*
 * Copyright 2001-2015 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Cells. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package com.aspose.cells.examples.files.utility;

import com.aspose.cells.Chart;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class ConvertChartToPdf {

    public static void main(String[] args)
            throws Exception {

        String dataDir = Utils.getDataDir(ConvertChartToPdf.class);
        String inputPath = dataDir + "Sample1.xls";
        String outputPath = dataDir + "Output-chart.pdf";

        //Load excel file containing charts
        Workbook workbook = new Workbook(inputPath);

        //Access first worksheet
        Worksheet worksheet = workbook.getWorksheets().get(0);

        //Access first chart inside the worksheet
        Chart chart = worksheet.getCharts().get(0);

        //Save the chart into pdf format
        chart.toPdf(outputPath);

        System.out.println("File saved " + outputPath);
    }
}

