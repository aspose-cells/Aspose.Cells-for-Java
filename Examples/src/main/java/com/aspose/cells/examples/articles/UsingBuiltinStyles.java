/*
 * Copyright 2001-2016 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Cells. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package com.aspose.cells.examples.articles;

import com.aspose.cells.BuiltinStyleType;
import com.aspose.cells.Cell;
import com.aspose.cells.Style;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class UsingBuiltinStyles {

    public static void main(String[] args)
            throws Exception {

        String dataDir = Utils.getDataDir(UsingBuiltinStyles.class);
        String output1Path = dataDir + "Output.xlsx";
        String output2Path = dataDir + "Output.ods";

        Workbook workbook = new Workbook();
        Style style = workbook.createBuiltinStyle(BuiltinStyleType.TITLE);

        Cell cell = workbook.getWorksheets().get(0).getCells().get("A1");
        cell.putValue("Aspose");
        cell.setStyle(style);

        workbook.getWorksheets().get(0).autoFitColumn(0);
        workbook.getWorksheets().get(0).autoFitRow(0);

        workbook.save(output1Path);
        System.out.println("File saved " + output1Path);
        workbook.save(output2Path);
        System.out.println("File saved " + output2Path);
    }
}

