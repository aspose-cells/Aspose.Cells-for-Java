package com.aspose.cells.examples.articles;

import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class ConvertTextNumericDatatoNumber {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ConvertTextNumericDatatoNumber.class) + "articles/";
		Workbook workbook = new Workbook(dataDir + "source.xlsx");

		for (int i = 0; i < workbook.getWorksheets().getCount(); i++) {
			workbook.getWorksheets().get(i).getCells().convertStringToNumericValue();
		}

		workbook.save(dataDir + "CTNDatatoNumber_out.xlsx");

	}
}
