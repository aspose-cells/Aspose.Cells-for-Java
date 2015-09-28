/* 
 * Copyright 2001-2015 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Cells. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package com.aspose.cells.examples.articles;

import com.aspose.cells.MetadataOptions;
import com.aspose.cells.MetadataType;
import com.aspose.cells.Workbook;
import com.aspose.cells.WorkbookMetadata;
import com.aspose.cells.examples.Utils;

public class UsingWorkbookMetadata {

    public static void main(String[] args) throws Exception {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(UsingWorkbookMetadata.class);


        // Open Workbook metadata
        MetadataOptions options = new MetadataOptions(MetadataType.DOCUMENT_PROPERTIES);
        WorkbookMetadata meta = new WorkbookMetadata(dataDir + "Sample1.xlsx", options);

        // Set some properties
        meta.getCustomDocumentProperties().add("test", "test");

        // Save the metadata info
        meta.save(dataDir + "Sample2.xlsx");

        // Open the workbook
        Workbook w = new Workbook(dataDir + "Sample2.xlsx");

        // Read document property
        System.out.println(w.getCustomDocumentProperties().get("test"));
    }
}
