/* 
 * Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Cells. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package com.aspose.cells.examples.data.handling.importing;

import com.aspose.cells.Cells;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class ImportingFromArray {

    public static void main(String[] args) throws Exception {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(ImportingFromArray.class);

        //Instantiating a Workbook object
        Workbook workbook = new Workbook();

        //Obtaining the reference of the worksheet
        Worksheet worksheet = workbook.getWorksheets().get(0);

        //Creating an array containing names as string values
        String[] names = new String[]{"laurence chen", "roman korchagin", "kyle huang"};

        //Importing the array of names to 1st row and first column vertically
        Cells cells = worksheet.getCells();
        cells.importArray(names, 0, 0, false);

        //Saving the Excel file
        workbook.save(dataDir + "DataImport.xls");

        //Printing the name of the cell found after searching worksheet
        System.out.println("Process completed successfully");

    }
}
