package AsposeCellsExamples.WorkbookSettings;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class ImplementErrorsAndBooleanValueInRussianOrAnyOtherLanguage { 

	// Russian Globalization
	class RussianGlobalization extends GlobalizationSettings {
		public String getErrorValueString(String err) {
			switch (err.toUpperCase()) {
			case "#NAME?":
				return "#RussianName-имя?";

			}

			return "RussianError-ошибка";
		}

		public String getBooleanValueString(Boolean bv) {
			return bv ? "RussianTrue-правда" : "RussianFalse-ложный";
		}
	}

	public void Run() throws Exception {
		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());

		String srcDir = Utils.Get_SourceDirectory();
		String outDir = Utils.Get_OutputDirectory();

		// Load the source workbook
		Workbook wb = new Workbook(srcDir + "sampleRussianGlobalization.xlsx");

		// Set GlobalizationSettings in Russian Language
		wb.getSettings().setGlobalizationSettings(new RussianGlobalization());

		// Calculate the formula
		wb.calculateFormula();

		// Save the workbook in pdf format
		wb.save(outDir + "outputRussianGlobalization.pdf");

		// Print the message
		System.out.println("ImplementErrorsAndBooleanValueInRussianOrAnyOtherLanguage executed successfully.");
	}

	public static void main(String[] args) throws Exception {

		ImplementErrorsAndBooleanValueInRussianOrAnyOtherLanguage impErr = new ImplementErrorsAndBooleanValueInRussianOrAnyOtherLanguage();
		impErr.Run();
	}
}
