package com.aspose.cells.examples.articles;

import com.aspose.cells.examples.Utils;
import com.aspose.cells.*;

public class FilterObjectsLoadingWorkbook {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(FilterObjectsLoadingWorkbook.class) + "articles/";

		//Filter charts from entire workbook
		LoadOptions ldOpts = new LoadOptions();
		ldOpts.setLoadFilter(new LoadFilter(LoadDataFilterOptions.ALL & ~LoadDataFilterOptions.CHART));

		//Load the workbook with above filter
		Workbook wb = new Workbook(dataDir + "sampleFilterCharts.xlsx", ldOpts);

		//Save entire worksheet into a single page
		PdfSaveOptions opts = new PdfSaveOptions();
		opts.setOnePagePerSheet(true);

		//Save the workbook in pdf format with the above pdf save options
		wb.save(dataDir + "sampleFilterCharts.pdf", opts);
	}
}