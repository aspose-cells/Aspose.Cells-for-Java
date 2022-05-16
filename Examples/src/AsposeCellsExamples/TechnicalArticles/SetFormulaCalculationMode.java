package AsposeCellsExamples.TechnicalArticles;

import com.aspose.cells.CalcModeType;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;

public class SetFormulaCalculationMode {
	public static void main(String[] args) throws Exception {
		// Create a workbook
		Workbook workbook = new Workbook();

		// Set the Formula Calculation Mode to Manual
		workbook.getSettings().getFormulaSettings().setCalculationMode(CalcModeType.MANUAL);

		// Save the workbook
		workbook.save("SFCalculationMode_out.xlsx", SaveFormat.XLSX);

	}
}
