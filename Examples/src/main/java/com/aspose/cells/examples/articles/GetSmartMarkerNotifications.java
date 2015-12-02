/*
 * Copyright 2001-2015 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Cells. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package com.aspose.cells.examples.articles;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;

public class GetSmartMarkerNotifications {

    public static void main(String[] args)
            throws Exception {

        String dataDir = Utils.getDataDir(GetSmartMarkerNotifications.class);
        String outputPath = dataDir + "Output.xlsx";

        //Instantiate a new Workbook designer
        WorkbookDesigner report = new WorkbookDesigner();

        //Get the first worksheet of the workbook
        Worksheet sheet = report.getWorkbook().getWorksheets().get(0);

        //Set the Variable Array marker to a cell
        //You may also place this Smart Marker into a template file manually using Excel and then open this file via WorkbookDesigner
        sheet.getCells().get("A1").putValue("&=$VariableArray");

        //Set the data source for the marker(s)
        report.setDataSource("VariableArray", new String[]{"English", "Arabic", "Hindi", "Urdu", "French"});

        //Set the CallBack property
        report.setCallBack(new SmartMarkerCallBack(report.getWorkbook()));

        //Process the markers
        report.process(false);

        //Save the result
        report.getWorkbook().save(outputPath);
        System.out.println("File saved " + outputPath);
    }
}

class SmartMarkerCallBack implements ISmartMarkerCallBack {
    Workbook workbook;

    SmartMarkerCallBack(Workbook workbook) {
        this.workbook = workbook;
    }

    @Override
    public void process(int sheetIndex, int rowIndex, int colIndex, String tableName, String columnName) {
        System.out.println("Processing Cell: " + workbook.getWorksheets().get(sheetIndex).getName() + "!" + CellsHelper.cellIndexToName(rowIndex, colIndex));
        System.out.println("Processing Marker: " + tableName + "." + columnName);
    }
}


