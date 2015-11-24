package com.aspose.cells.examples.featurescomparison.workbook.adddataincells;

import com.aspose.cells.Cell;
import com.aspose.cells.Cells;
import com.aspose.cells.Style;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class AsposeInsertCellsData
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(AsposeInsertCellsData.class);

        //Instantiating a Workbook object
        Workbook workbook = new Workbook();

        //Accessing the added worksheet in the Excel file
        Worksheet worksheet = workbook.getWorksheets().get(0);
        Cells cells = worksheet.getCells();

        cells.get("A1").setValue("Hello World"); //Adding a string value to the cell
        cells.get("A2").setValue(20.5); //Adding a double value to the cell
        cells.get("A3").setValue(15); //Adding an integer  value to the cell
        cells.get("A4").setValue(true); //Adding a boolean value to the cell

        Cell cell = cells.get("A5"); //Adding a date/time value to the cell
        cell.setValue(java.util.Calendar.getInstance());

        //Setting the display format of the date
        Style style = cell.getStyle();
        style.setNumber(15);
        cell.setStyle(style);

        workbook.save(dataDir + "DataInCells_Aspose_Out.xls"); //Saving the Excel file

        // Print message
        System.out.println("Data Added Successfully");
    }
}
