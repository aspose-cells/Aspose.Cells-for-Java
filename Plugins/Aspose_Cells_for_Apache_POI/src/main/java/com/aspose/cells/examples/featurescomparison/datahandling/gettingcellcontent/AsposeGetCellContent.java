package com.aspose.cells.examples.featurescomparison.datahandling.gettingcellcontent;

import com.aspose.cells.Cell;
import com.aspose.cells.CellValueType;
import com.aspose.cells.Cells;
import com.aspose.cells.Range;
import com.aspose.cells.RowCollection;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class AsposeGetCellContent
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(AsposeGetCellContent.class);

        Workbook workbook = new Workbook(dataDir + "workbook.xls");

        //Accessing the first worksheet in the Excel file
        Worksheet worksheet = workbook.getWorksheets().get(0);
        Cells cells = worksheet.getCells();

        //Access the Maximum Display Range
        Range range = worksheet.getCells().getMaxDisplayRange();
        int tcols = range.getColumnCount();
        int trows = range.getRowCount();

        System.out.println("Total Rows:" + trows);
        System.out.println("Total Cols:" + tcols);

        // Access value of Cell B4
        //=====================================================
        System.out.println(cells.get("B4").getValue());

        Cell cell = cells.get(3,1); //Access value of Cell B4
        System.out.println(cell.getValue());
        //=====================================================
        RowCollection rows = cells.getRows();

        for (int i = 0 ; i < rows.getCount() ; i++)
        {
            for (int j = 0 ; j < tcols ; j++)
            {
                if (cells.get(i,j).getType() != CellValueType.IS_NULL)
                {
                        System.out.println(cells.get(i,j).getName() + " - " + cells.get(i,j).getValue());
                }
            }
        }
    }
}