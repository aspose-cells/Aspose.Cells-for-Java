package AsposeCellsExamples.TechnicalArticles;

import com.aspose.cells.CalcModeType;
import com.aspose.cells.CalculationOptions;
import com.aspose.cells.Cell;
import com.aspose.cells.Cells;
import com.aspose.cells.Style;
import com.aspose.cells.Workbook;
import AsposeCellsExamples.Utils;

public class ReturningRangeOfValues {
	public static void main(String[] args) throws Exception {
		// ExStart:1
		String outputDir = Utils.Get_OutputDirectory();
		Workbook workbook = new Workbook();
		Cells cells = workbook.getWorksheets().get(0).getCells();

		Cell cell = cells.get(0, 0);
		cell.setArrayFormula("=MYFUNC()", 2, 2);

		Style style = cell.getStyle();
		style.setNumber(14);
		cell.setStyle(style);

		CalculationOptions copt = new CalculationOptions();
		copt.setCustomEngine(new CustomFunctionStaticValue());
		workbook.calculateFormula(copt);

		// Save to XLSX by setting the calc mode to manual
		workbook.getSettings().setCalcMode(CalcModeType.MANUAL);
		workbook.save(outputDir + "output.xlsx");

		// Save to PDF
		workbook.save(outputDir + "output.pdf");
		// ExEnd:1
	}
}
