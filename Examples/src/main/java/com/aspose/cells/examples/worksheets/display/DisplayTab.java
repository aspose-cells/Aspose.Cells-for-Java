package com.aspose.cells.examples.worksheets.display;

import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class DisplayTab {

    public static void main(String[] args) throws Exception {
        //ExStart:1
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(DisplayTab.class);

        //Instantiating a Workbook object by excel file path
        Workbook workbook = new Workbook(dataDir + "book1.xls");

        //Hiding the tabs of the Excel file
        workbook.getSettings().setShowTabs(true);

        //Saving the modified Excel file in default (that is Excel 2003) format
        workbook.save(dataDir + "output.xls");

        // Print message
        System.out.println("Tabs are now displayed, please check the output file.");
        //ExEnd:1
    }
}
