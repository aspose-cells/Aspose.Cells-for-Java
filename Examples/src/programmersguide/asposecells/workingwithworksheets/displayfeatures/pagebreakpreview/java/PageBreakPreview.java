/* 
 * Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Cells. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
 
package programmersguide.asposecells.workingwithworksheets.displayfeatures.pagebreakpreview.java;

import com.aspose.cells.*;

public class PageBreakPreview
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = "src/programmersguide/asposecells/workingwithworksheets/displayfeatures/pagebreakpreview/data/";

        //Instantiating a Excel object by excel file path
        Workbook workbook = new Workbook(dataDir + "book1.xls");

        //Adding a new worksheet to the Workbook object
        WorksheetCollection worksheets = workbook.getWorksheets();

        Worksheet worksheet = worksheets.get(0);

        //Displaying the worksheet in page break preview
        worksheet.setPageBreakPreview(true);

        //Saving the modified Excel file in default format
        workbook.save(dataDir + "output.xls");

        // Print message
        System.out.println("Page break preview is enabled for sheet 1, please check the output document.");
    }
}




