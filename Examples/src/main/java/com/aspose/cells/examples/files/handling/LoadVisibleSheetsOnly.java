/*
 * Copyright 2001-2016 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Cells. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package com.aspose.cells.examples.files.handling;

import com.aspose.cells.LoadDataOption;
import com.aspose.cells.LoadOptions;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class LoadVisibleSheetsOnly {

    public static void main(String[] args) throws Exception {
        String dataDir = Utils.getDataDir(LoadVisibleSheetsOnly.class);
        String sampleFile = "Sample.out.xlsx";
        String samplePath = dataDir + sampleFile;

        // Create a sample workbook
        // and put some data in first cell of all 3 sheets
        Workbook createWorkbook = new Workbook();
        createWorkbook.getWorksheets().get("Sheet1").getCells().get("A1").setValue("Aspose");
        createWorkbook.getWorksheets().add("Sheet2").getCells().get("A1").setValue("Aspose");
        createWorkbook.getWorksheets().add("Sheet3").getCells().get("A1").setValue("Aspose");
        // Keep Sheet3 invisible
        createWorkbook.getWorksheets().get("Sheet3").setVisible(false);
        createWorkbook.save(samplePath);

        // Load the sample workbook
        LoadDataOption loadDataOption = new LoadDataOption();
        loadDataOption.setOnlyVisibleWorksheet(true);
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.setLoadDataAndFormatting(true);
        loadOptions.setLoadDataOptions(loadDataOption);

        Workbook loadWorkbook = new Workbook(samplePath, loadOptions);

        System.out.println("Sheet1: A1: " + loadWorkbook.getWorksheets().get("Sheet1").getCells().get("A1").getValue());
        System.out.println("Sheet2: A1: " + loadWorkbook.getWorksheets().get("Sheet2").getCells().get("A1").getValue());
        System.out.println("Sheet3: A1: " + loadWorkbook.getWorksheets().get("Sheet3").getCells().get("A1").getValue());

        System.out.println("Data is not loaded from invisible sheet");
    }
}

