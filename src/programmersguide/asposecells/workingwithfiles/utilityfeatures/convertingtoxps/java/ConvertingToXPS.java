/* 
 * Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Cells. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
 
package programmersguide.asposecells.workingwithfiles.utilityfeatures.convertingtoxps.java;

import com.aspose.cells.*;

public class ConvertingToXPS
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = "src/programmersguide/asposecells/workingwithfiles/utilityfeatures/convertingtoxps/data/";

        Workbook workbook = new Workbook(dataDir + "Book1.xls");

        //Get the first worksheet.
        Worksheet sheet = workbook.getWorksheets().get(0);

        //Apply different Image and Print options
        com.aspose.cells.ImageOrPrintOptions options = new ImageOrPrintOptions();

        //Set the Format
        options.setSaveFormat(SaveFormat.XPS);

        // Render the sheet with respect to specified printing options
        com.aspose.cells.SheetRender sr = new SheetRender(sheet, options);
        sr.toImage(0, dataDir + "out_printingxps.xps");

        //Save the complete Workbook in XPS format
        workbook.save(dataDir + "out_whole_printingxps.xps", SaveFormat.XPS);

        // Print message
        System.out.println("Excel to XPS conversion performed successfully.");
    }
}




