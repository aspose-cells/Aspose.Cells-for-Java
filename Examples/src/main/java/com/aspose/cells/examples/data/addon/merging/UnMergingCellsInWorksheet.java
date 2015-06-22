/* 
 * Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Cells. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package com.aspose.cells.examples.data.addon.merging;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;

public class UnMergingCellsInWorksheet {

    public static void main(String[] args) throws Exception {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(UnMergingCellsInWorksheet.class);

        //Create a Workbook.
        Workbook wbk = new Workbook(dataDir + "mergingcells.xls");

        //Create a Worksheet and get the first sheet.
        Worksheet worksheet = wbk.getWorksheets().get(0);

        //Create a Cells object to fetch all the cells.
        Cells cells = worksheet.getCells();

        //Unmerge the cells.
        cells.unMerge(5, 2, 2, 3);

        //Save the file.
        wbk.save(dataDir + "unmergingcells.xls");

        // Print message
        System.out.println("Process completed successfully");
    }
}
