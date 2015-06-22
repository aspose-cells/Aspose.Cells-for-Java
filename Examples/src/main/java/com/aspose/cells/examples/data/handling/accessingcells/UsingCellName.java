/* 
 * Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Cells. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package com.aspose.cells.examples.data.handling.accessingcells;

import com.aspose.cells.Workbook;

import com.aspose.cells.examples.Utils;

public class UsingCellName {

    public static void main(String[] args) throws Exception {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(UsingCellName.class);

        //Instantiating a Workbook object
        Workbook workbook = new Workbook(dataDir + "book1.xls");

        //Accessing the worksheet in the Excel file
        com.aspose.cells.Worksheet worksheet = workbook.getWorksheets().get(0);
        com.aspose.cells.Cells cells = worksheet.getCells();

        //Accessing a cell using its name
        com.aspose.cells.Cell cell = cells.get("A1");

        // Print message
        System.out.println("Cell Value: " + cell.getValue());
    }
}
