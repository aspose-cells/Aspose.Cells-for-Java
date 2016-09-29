package com.aspose.cells.examples.articles;

import com.aspose.cells.CopyOptions;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class ChangeDataSource {
	public static void main(String[] args) throws Exception {

		String dataDir = Utils.getSharedDataDir(ChangeDataSource.class) + "articles/";
		// Load sample excel file
		Workbook wb = new Workbook(dataDir + "sample.xlsx");

		// Access the first sheet which contains chart
		Worksheet source = wb.getWorksheets().get(0);

		// Add another sheet named DestSheet
		Worksheet destination = wb.getWorksheets().add("DestSheet");

		// Set CopyOptions.ReferToDestinationSheet to true
		CopyOptions options = new CopyOptions();
		options.setReferToDestinationSheet(true);

		/*
		 * Copy all the rows of source worksheet to destination worksheet which includes chart as well The chart data source will
		 * now refer to DestSheet
		 */
		destination.getCells().copyRows(source.getCells(), 0, 0, source.getCells().getMaxDisplayRange().getRowCount(),
				options);

		// Save workbook in xlsx format
		wb.save(dataDir + "CDataSource_out.xlsx", SaveFormat.XLSX);

	}
}
