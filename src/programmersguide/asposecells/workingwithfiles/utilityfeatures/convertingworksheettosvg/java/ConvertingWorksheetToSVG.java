/* 
 * Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Cells. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
 
package programmersguide.asposecells.workingwithfiles.utilityfeatures.convertingworksheettosvg.java;

import com.aspose.cells.*;

public class ConvertingWorksheetToSVG
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = "src/programmersguide/asposecells/workingwithfiles/utilityfeatures/convertingworksheettosvg/data/";

        String path = dataDir + "Template.xlsx";

        //Create a workbook object from the template file
        Workbook workbook = new Workbook(path);

        //Convert each worksheet into svg format in a single page.
        ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();
        imgOptions.setSaveFormat(SaveFormat.SVG);
        imgOptions.setOnePagePerSheet(true);

        //Convert each worksheet into svg format
        int sheetCount = workbook.getWorksheets().getCount();

        for(int i=0; i<sheetCount; i++)
        {
            Worksheet sheet = workbook.getWorksheets().get(i);

            SheetRender sr = new SheetRender(sheet, imgOptions);

            for (int k = 0; k < sr.getPageCount(); k++)
            {
                //Output the worksheet into Svg image format
                sr.toImage(k, path + sheet.getName() + k + ".out.svg");
            }
        }

        // Print message
        System.out.println("Excel to SVG conversion completed successfully.");
    }
}




