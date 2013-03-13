/* 
 * Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Cells. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
 
package programmersguide.asposecells.workingwithworksheets.displayfeatures.displayhidetabs.java;

import com.aspose.cells.*;

public class DisplayHideTabs
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = "src/programmersguide/asposecells/workingwithworksheets/displayfeatures/displayhidetabs/data/";

        //Instantiating a Workbook object by excel file path
        Workbook workbook = new Workbook(dataDir + "book1.xls");

        //Hiding the tabs of the Excel file
        workbook.getSettings().setShowTabs(false);

        //Saving the modified Excel file in default (that is Excel 2003) format
        workbook.save(dataDir + "output.xls");

        // Print message
        System.out.println("Tabs are now hidden, please check the output file.");
    }
}




