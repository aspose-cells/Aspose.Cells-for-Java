package AsposeCellsExamples.Formulas;

import com.aspose.cells.Cells;
import com.aspose.cells.DateTime;
import com.aspose.cells.Workbook;

import AsposeCellsExamples.Utils;

public class CalculatingFormulasOnce {
    public static void main(String[] args) throws Exception {
        // The path to the documents directory.
        String dataDir = Utils.getSharedDataDir(CalculatingFormulasOnce.class) + "formulas/";
        // Load the template workbook
        Workbook workbook = new Workbook(dataDir + "book1.xls");

        // Print the time before formula calculation
        System.out.println(DateTime.getNow());

        // Set the CreateCalcChain as true
        workbook.getSettings().setCreateCalcChain(true);

        // Calculate the workbook formulas
        workbook.calculateFormula();

        Cells cells = workbook.getWorksheets().get("Sheet1").getCells();
        
        //with original values, the calculated result
        System.out.println(cells.get("A11").getValue());
        
        //update one value the formula depends on
        cells.get("A5").putValue(15);
        
        // Calculate the workbook formulas again, in fact only A11 needs to be and will be calculated
        workbook.calculateFormula();

        //check the re-calculated value
        System.out.println(cells.get("A11").getValue());
        
        // Print the time after formula calculation
        System.out.println(DateTime.getNow());
    }
}
