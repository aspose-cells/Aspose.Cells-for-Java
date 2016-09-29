package com.aspose.cells.examples.articles;

import com.aspose.cells.LoadDataFilterOptions;
import com.aspose.cells.LoadFormat;
import com.aspose.cells.LoadOptions;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class FilterDataWhileLoadingWorkbook {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(FilterDataWhileLoadingWorkbook.class) + "articles/";
		// Set the load options, we only want to load shapes and do not want to load data
		LoadOptions opts = new LoadOptions(LoadFormat.XLSX);
		opts.setLoadDataFilterOptions(LoadDataFilterOptions.SHAPE);

		// Create workbook object from sample excel file using load options
		Workbook wb = new Workbook(dataDir + "sample.xlsx", opts);

		// Save the output in pdf format
		wb.save(dataDir + "FDWLWorkbook_out.pdf", SaveFormat.PDF);

	}

}
