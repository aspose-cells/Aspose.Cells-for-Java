/* 
 * Copyright 2001-2015 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Cells. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package com.aspose.cells.examples.articles;

import com.aspose.cells.ExternalConnection;
import com.aspose.cells.WebQueryConnection;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class GetDataConnection {

    public static void main(String[] args)
            throws Exception {

        String dataDir = Utils.getDataDir(GetDataConnection.class);
        String inputPath = dataDir + "WebQuerySample.xlsx";

        Workbook workbook = new Workbook(inputPath);

        ExternalConnection connection = workbook.getDataConnections().get(0);

        if (connection instanceof WebQueryConnection) {
            WebQueryConnection webQuery = (WebQueryConnection) connection;
            System.out.println("Web Query URL: " + webQuery.getUrl());
        }
    }
}
