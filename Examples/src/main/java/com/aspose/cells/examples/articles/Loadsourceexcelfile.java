package com.aspose.cells.examples.articles;

import com.aspose.cells.LoadDataFilterOptions;
import com.aspose.cells.LoadOptions;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class Loadsourceexcelfile {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(Loadsourceexcelfile.class) + "articles/";
		// Specify the load options and filter the data we do not want to load charts
		LoadOptions options = new LoadOptions();
		options.setLoadDataFilterOptions(LoadDataFilterOptions.ALL & ~LoadDataFilterOptions.CHART);

		// Load the workbook with specified load options
		Workbook workbook = new Workbook(dataDir + "sample.xlsx", options);

		// Save the workbook in output format
		workbook.save(dataDir + "LSourceexcelfile_out.pdf", SaveFormat.PDF);

	}

}
