package com.aspose.cells.examples.worksheets.management;

import java.io.FileInputStream;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;

public class RemovingWorksheetsusingSheetIndex {

    public static void main(String[] args) throws Exception {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(RemovingWorksheetsusingSheetIndex.class);

        //Creating a file stream containing the Excel file to be opened
        FileInputStream fstream = new FileInputStream(dataDir + "book.xls");

        //Instantiating a Workbook object with the stream
        Workbook workbook = new Workbook(fstream);

        //Removing a worksheet using its sheet index
        workbook.getWorksheets().removeAt(0);

        //Saving the Excel file
        workbook.save(dataDir + "book.out.xls");

        //Closing the file stream to free all resources
        fstream.close();

        //Print Message
        System.out.println("Sheet removed successfully.");
    }
}
