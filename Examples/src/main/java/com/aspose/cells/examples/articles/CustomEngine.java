package com.aspose.cells.examples.articles;

import com.aspose.cells.AbstractCalculationEngine;
import com.aspose.cells.CalculationData;
import com.aspose.cells.*;

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