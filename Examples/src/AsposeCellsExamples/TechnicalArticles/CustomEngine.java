package AsposeCellsExamples.TechnicalArticles;

import com.aspose.cells.AbstractCalculationEngine;
import com.aspose.cells.CalculationData;

public class CustomEngine extends AbstractCalculationEngine
{
	public void calculate(CalculationData data)
        {
		if(data.getFunctionName().toUpperCase().equals("SUM")==true)
                {
                    double val = (double)data.getCalculatedValue();
                    val = val + 30;

                    data.setCalculatedValue(val);
                }
        }
}