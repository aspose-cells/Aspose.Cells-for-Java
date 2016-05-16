package com.aspose.cells.examples.files.handling;

import com.aspose.cells.FileFormatType;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class SaveInSpreadsheetMLFormat {

    public static void main(String[] args) throws Exception {
        //ExStart:1
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(SaveInSpreadsheetMLFormat.class);

        //Creating an Workbook object with an Excel file path
        Workbook workbook = new Workbook();

        //Save in SpreadsheetML format
        workbook.save(dataDir + "output.xml", FileFormatType.EXCEL_2003_XML);

        //Print Message
        System.out.println("Worksheets are saved successfully.");
        //ExEnd:1
    }
}
