/* 
 * Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Cells. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package com.aspose.cells.examples.data.processing;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;

public class CreatingSubtotals {

    public static void main(String[] args) throws Exception {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(CreatingSubtotals.class);

        //Instantiate a new workbook
        Workbook workbook = new Workbook(dataDir + "book1.xls");

        //Get the Cells collection in the first worksheet
        Cells cells = workbook.getWorksheets().get(0).getCells();

        //Create a cellarea i.e.., B3:C19
        CellArea ca = new CellArea();
        ca.StartRow = 2;
        ca.StartColumn = 1;
        ca.EndRow = 18;
        ca.EndColumn = 2;

        //Apply subtotal, the consolidation function is Sum and it will applied to
        //Second column (C) in the list
        cells.subtotal(ca, 0, ConsolidationFunction.SUM, new int[]{1});

        //Save the excel file
        workbook.save(dataDir + "output.xls");

        // Print message
        System.out.println("Process completed successfully");

    }
}
