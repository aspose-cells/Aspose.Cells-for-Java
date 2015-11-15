/* 
 * Copyright 2001-2015 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Cells. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package com.aspose.cells.examples.articles;

import com.aspose.cells.HtmlLinkTargetType;
import com.aspose.cells.HtmlSaveOptions;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class ChangeHtmlLinkTarget {

    public static void main(String[] args)
            throws Exception {

        String dataDir = Utils.getDataDir(ChangeHtmlLinkTarget.class);
        String inputPath = dataDir + "Sample1.xlsx";
        String outputPath = dataDir + "Output.html";

        Workbook workbook = new Workbook(inputPath);

        HtmlSaveOptions opts = new HtmlSaveOptions();
        opts.setLinkTargetType(HtmlLinkTargetType.SELF);

        workbook.save(outputPath, opts);

        System.out.println("File saved " + outputPath);
    }
}
