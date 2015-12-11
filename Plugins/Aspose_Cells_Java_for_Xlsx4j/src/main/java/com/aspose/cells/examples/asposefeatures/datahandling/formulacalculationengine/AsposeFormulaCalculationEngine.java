package com.aspose.cells.examples.asposefeatures.datahandling.formulacalculationengine;

import com.aspose.cells.Cell;
import com.aspose.cells.Cells;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class AsposeFormulaCalculationEngine
{
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(AsposeFormulaCalculationEngine.class);

        //Instantiating a Workbook object
        Workbook book = new Workbook();

        //Obtaining the reference of the newly added worksheet
        int sheetIndex = book.getWorksheets().add();
        Worksheet worksheet = book.getWorksheets().get(sheetIndex);
        Cells cells = worksheet.getCells();
        Cell cell = null;

        //Adding a value to "A1" cell
        cell = cells.get("A1");
        cell.setValue(1);

        //Adding a value to "A2" cell
        cell = cells.get("A2");
        cell.setValue(2);

        //Adding a value to "A3" cell
        cell = cells.get("A3");
        cell.setValue(3);

        //Adding a SUM formula to "A4" cell
        cell = cells.get("A4");
        cell.setFormula("=SUM(A1:A3)");

        //Calculating the results of formulas
        book.calculateFormula();

        //Saving the Excel file
        book.save(dataDir + "FormulaEngine_Out.xls");
    }
}
