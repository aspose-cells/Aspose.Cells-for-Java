package com.aspose.cells.examples.charts.ManipulateChart;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;

public class MicrosoftTheme {

    public static void main(String[] args) throws Exception {
        //ExStart:1
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(MicrosoftTheme.class);

        //Instantiating a Workbook object
        Workbook workbook = new Workbook(dataDir + "Book1.xlsx");

        //Obtaining the reference of the first worksheet
        WorksheetCollection worksheets = workbook.getWorksheets();
        Worksheet sheet = worksheets.get(0);

        Chart chart =sheet.getCharts().get(0);

        //Specify the FilFormat's type to Solid Fill of the first series
        chart.getNSeries().get(0).getArea().getFillFormat().setType(FillType.SOLID);

        //Get the CellsColor of SolidFill
        CellsColor cc = chart.getNSeries().get(0).getArea().getFillFormat().getSolidFill().getCellsColor();

        //Create a theme in Accent style
        cc.setThemeColor(new ThemeColor(ThemeColorType.FOLLOWED_HYPERLINK, 0.6));

        //Apply the them to the series
        chart.getNSeries().get(0).getArea().getFillFormat().getSolidFill().setCellsColor(cc);

        //Save the Excel file
        workbook.save(dataDir + "output.xlsx");

        // Print message
        System.out.println("MicrosoftTheme is successfully applied.");
        //ExEnd:1
    }
}
