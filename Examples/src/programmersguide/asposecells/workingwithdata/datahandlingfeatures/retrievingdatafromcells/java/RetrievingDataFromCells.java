/* 
 * Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Cells. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
 
package programmersguide.asposecells.workingwithdata.datahandlingfeatures.retrievingdatafromcells.java;

import com.aspose.cells.Workbook;

public class RetrievingDataFromCells
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = "src/programmersguide/asposecells/workingwithdata/datahandlingfeatures/retrievingdatafromcells/data/";

        //Instantiating a Workbook object
        Workbook workbook = new Workbook();

        //Accessing the worksheet
        com.aspose.cells.Worksheet worksheet = workbook.getWorksheets().get(0);
        com.aspose.cells.Cells cells = worksheet.getCells();

        //get cell from cells collection
        com.aspose.cells.Cell cell = cells.get("A5");

        switch(cell.getType())
        {
            case com.aspose.cells.CellValueType.IS_BOOL:
                System.out.println("Boolean Value: " + cell.getValue());
                break;
            case com.aspose.cells.CellValueType.IS_DATE_TIME:
                System.out.println("Date Value: " + cell.getValue())  ;
                break;
            case com.aspose.cells.CellValueType.IS_NUMERIC:
                System.out.println("Numeric Value: " + cell.getValue())   ;
                break;
            case com.aspose.cells.CellValueType.IS_STRING:
                System.out.println("String Value: " + cell.getValue())     ;
                break;
            case com.aspose.cells.CellValueType.IS_NULL:
                System.out.println("Null Value");
                break;
        }

    }
}




