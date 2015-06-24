/* 
 * Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Cells. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package com.aspose.cells.examples.data.addon.namedranges;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;

public class RenameNamedRange {

    public static void main(String[] args) throws Exception {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(RenameNamedRange.class);

        //Open an existing Excel file that has a (global) named range "TestRange" in it
        Workbook workbook = new Workbook(dataDir + "book1.xls");

        //Get the first worksheet
        Worksheet sheet = workbook.getWorksheets().get(0);

        //Get the Cells of the sheet
        Cells cells = sheet.getCells();

        //Get the named range "MyRange"
        Name name = workbook.getWorksheets().getNames().get("TestRange");

        //Rename it
        name.setText("NewRange");

        //Save the Excel file
        workbook.save(dataDir + "RenamingRange.xlsx");

        // Print message
        System.out.println("Process completed successfully");
    }
}
