package com.aspose.cells.examples.rowscolumns;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;

public class AutoFitRowsandColumns {

    public static void main(String[] args) throws Exception {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(AutoFitRowsandColumns.class);

        //Instantiating a Workbook object
        Workbook workbook = new Workbook(dataDir + "workbook.xls");

        //Accessing the first worksheet in the Excel file
        Worksheet worksheet = workbook.getWorksheets().get(0);

        //Auto-fitting the 2nd row of the worksheet
        worksheet.autoFitRow(1);

        //Auto-fitting the 1st column of the worksheet
        worksheet.autoFitColumn(0);

        //Saving the modified Excel file in default (that is Excel 2003) format
        workbook.save(dataDir + "workbook.out.xls");

        //Print message
        System.out.println("Row and Column auto fit successfully.");
    }
}
