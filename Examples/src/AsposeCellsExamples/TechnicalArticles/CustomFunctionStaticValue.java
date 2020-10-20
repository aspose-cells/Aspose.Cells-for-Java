package AsposeCellsExamples.TechnicalArticles;

//ExStart:CustomFunctionStaticValue
import java.util.ArrayList;

import com.aspose.cells.AbstractCalculationEngine;
import com.aspose.cells.CalculationData;
import com.aspose.cells.DateTime;

public class CustomFunctionStaticValue extends AbstractCalculationEngine {
	@Override
	public void calculate(CalculationData calculationData) {
		calculationData.setCalculatedValue(new Object[][] { new Object[] { new DateTime(2015, 6, 12, 10, 6, 30), 2 },
				new Object[] { 3.0, "Test" } });
	}
}
//ExEnd:CustomFunctionStaticValue
