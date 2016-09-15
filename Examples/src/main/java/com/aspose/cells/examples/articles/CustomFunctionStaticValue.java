package com.aspose.cells.examples.articles;

//ExStart:CustomFunctionStaticValue
import java.util.ArrayList;

import com.aspose.cells.DateTime;
import com.aspose.cells.ICustomFunction;

public class CustomFunctionStaticValue implements ICustomFunction {

	@Override
	public Object calculateCustomFunction(String functionName, ArrayList paramsList, ArrayList contextObjects) {
		return new Object[][] { new Object[] { new DateTime(2015, 6, 12, 10, 6, 30), 2 },
				new Object[] { 3.0, "Test" } };
	}

}
