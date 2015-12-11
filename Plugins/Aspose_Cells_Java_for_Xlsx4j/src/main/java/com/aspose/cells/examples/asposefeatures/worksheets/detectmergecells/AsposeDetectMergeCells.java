package com.aspose.cells.examples.asposefeatures.worksheets.detectmergecells;

import java.util.ArrayList;

import com.aspose.cells.CellArea;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class AsposeDetectMergeCells
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(AsposeDetectMergeCells.class);

        //Instantiate a new Workbook
        Workbook wkBook = new Workbook(dataDir + "MergeInput.xls");

        //Get a worksheet in the workbook
        Worksheet wkSheet = wkBook.getWorksheets().get(0);

        //Clear its contents
        wkSheet.getCells().clearContents(0,0,wkSheet.getCells().getMaxDataRow(),wkSheet.getCells().getMaxDataColumn());

        //Create an arraylist object
        //Get the merged cells list to put it into the arraylist object       
        ArrayList<CellArea> al = wkSheet.getCells().getMergedCells();

        //Define cellarea
        CellArea ca;

        //Define some variables
        int frow, fcol, erow, ecol;

        // Print Message
        System.out.println("Merged Areas: \n"+ al.toString());

        //Loop through the arraylist and get each cellarea to unmerge it
        for(int i = al.size()-1 ; i > -1; i--)
        { 
                ca = new CellArea();
                ca = (CellArea)al.get(i);
                frow = ca.StartRow;
                fcol = ca.StartColumn;
                erow = ca.EndRow;
                ecol = ca.EndColumn;
                System.out.println((i+1) + ". [" + fcol +"," + frow +"] " + "[" + ecol +"," + erow +"]");
                wkSheet.getCells().unMerge(frow, fcol, erow, ecol);
        }

        //Save the excel file
        wkBook.save(dataDir + "AsposeMergeOutput.xls");

        // Print Message
        System.out.println("Detect Merge Cells successful.");
    }
}
