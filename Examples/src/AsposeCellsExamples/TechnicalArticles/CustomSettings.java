package AsposeCellsExamples.TechnicalArticles;

import com.aspose.cells.ConsolidationFunction;
import com.aspose.cells.GlobalizationSettings;

public class CustomSettings extends GlobalizationSettings {
	public String getTotalName(int functionType) {
		switch (functionType) {
		case ConsolidationFunction.AVERAGE:
			return "AVG";

		// Handle other cases

		default:
			return super.getTotalName(functionType);
		}
	}

	public String getGrandTotalName(int functionType) {
		switch (functionType) {
		case ConsolidationFunction.AVERAGE:
			return "GRAND AVG";

		// Handle other cases

		default:
			return super.getGrandTotalName(functionType);
		}
	}
}
