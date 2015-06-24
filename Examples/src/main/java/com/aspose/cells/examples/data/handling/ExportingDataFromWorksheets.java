/* 
 * Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Cells. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package com.aspose.cells.examples.data.handling;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;
import java.io.*;

public class ExportingDataFromWorksheets {

    public static void main(String[] args) throws Exception {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(ExportingDataFromWorksheets.class);

        //Creating a file stream containing the Excel file to be opened
        FileInputStream fstream = new FileInputStream(dataDir + "book1.xls");

        //Instantiating a Workbook object
        Workbook workbook = new Workbook(fstream);

        //Accessing the first worksheet in the Excel file
        Worksheet worksheet = workbook.getWorksheets().get(0);

        //Exporting the contents of 7 rows and 2 columns starting from 1st cell to Array.
        Object dataTable[][] = worksheet.getCells().exportArray(0, 0, 7, 2);

        //Printing the name of the cell found after searching worksheet
        System.out.println("No. Of Rows Imported: " + dataTable.length);

        //Closing the file stream to free all resources
        fstream.close();

    }
}
