/* 
 * Copyright 2001-2015 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Cells. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package com.aspose.cells.examples.articles;

import com.aspose.cells.HtmlSaveOptions;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class HtmlExportFrameScripts {

    public static void main(String[] args) throws Exception {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(HtmlExportFrameScripts.class);

        // Open the required workbook to convert
        Workbook w = new Workbook(dataDir + "Sample1.xlsx");

        // Disable exporting frame scripts and document properties
        HtmlSaveOptions options = new HtmlSaveOptions();
        options.setExportFrameScriptsAndProperties(false);

        // Save workbook as HTML
        w.save(dataDir + "output.html", options);

        System.out.println("File saved");
    }
}
