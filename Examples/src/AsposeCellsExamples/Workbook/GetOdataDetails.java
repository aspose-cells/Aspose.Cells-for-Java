package AsposeCellsExamples.Workbook;

import AsposeCellsExamples.Utils;
import com.aspose.cells.*;

public class GetOdataDetails {

	public static void main(String[] args) throws Exception {

        // ExStart:1
        // The path to the directories.
        String sourceDir = Utils.Get_SourceDirectory();

        Workbook workbook = new Workbook(sourceDir + "ODataSample.xlsx");
        PowerQueryFormulaCollection PQFcoll = workbook.getDataMashup().getPowerQueryFormulas();
        for (Object obj : PQFcoll)
        {
            PowerQueryFormula PQF = (PowerQueryFormula)obj;
            System.out.println("Connection Name: " + PQF.getName());
            PowerQueryFormulaItemCollection PQFIcoll = PQF.getPowerQueryFormulaItems();
            for (Object obj2 : PQFIcoll)
            {
                PowerQueryFormulaItem PQFI = (PowerQueryFormulaItem)obj2;
                System.out.println("Name: " + PQFI.getName());
                System.out.println("Value: " + PQFI.getValue());
            }
        }
        // ExEnd:1

		System.out.println("GetOdataDetails executed successfully.");
	}
}
