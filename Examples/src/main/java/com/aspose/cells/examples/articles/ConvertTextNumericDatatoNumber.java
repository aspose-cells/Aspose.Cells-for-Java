package com.aspose.cells.examples.articles;

import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class ConvertTextNumericDatatoNumber {
	public static void main(String[] args) throws Exception {
		// ExStart:ConvertTextNumericDatatoNumber
		// The path to the documents directory.
		String dataDir = Utils.getDataDir(ConvertTextNumericDatatoNumber.class);
		Workbook workbook = new Workbook(dataDir + "source.xlsx");

		for (int i = 0; i < workbook.getWorksheets().getCount(); i++) {
			workbook.getWorksheets().get(i).getCells().convertStringToNumericValue();
		}

		workbook.save(dataDir + "output.xlsx");
		// ExEnd:ConvertTextNumericDatatoNumber
	}
}
