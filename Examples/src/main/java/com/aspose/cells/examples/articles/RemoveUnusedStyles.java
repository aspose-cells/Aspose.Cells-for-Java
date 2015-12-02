/* 
 * Copyright 2001-2015 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Cells. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package com.aspose.cells.examples.articles;

import com.aspose.cells.VbaProject;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class RemoveUnusedStyles {

    public static void main(String[] args)
            throws Exception {

        String dataDir = Utils.getDataDir(RemoveUnusedStyles.class);
        String inputPath = dataDir + "Styles.xlsx";
        String outputPath = dataDir + "Output.xlsx";

        Workbook workbook = new Workbook(inputPath);

        workbook.removeUnusedStyles();

        workbook.save(outputPath);
        System.out.println("File saved " + outputPath);
    }
}

