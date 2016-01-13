package com.aspose.cells.examples.data.handling.accessingcells;

import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class UsingCellIndexInCellsCollection {

    public static void main(String[] args) throws Exception {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(UsingCellIndexInCellsCollection.class);

        //Instantiating a Workbook object
        Workbook workbook = new Workbook(dataDir + "book1.xls");

        //Accessing the worksheet in the Excel file
        com.aspose.cells.Worksheet worksheet = workbook.getWorksheets().get(0);
        com.aspose.cells.Cells cells = worksheet.getCells();

        //Accessing a cell using cell index
        com.aspose.cells.Cell cell = cells.get(0, 0);

        // Print message
        System.out.println("Cell Value: " + cell.getValue());
    }
}
