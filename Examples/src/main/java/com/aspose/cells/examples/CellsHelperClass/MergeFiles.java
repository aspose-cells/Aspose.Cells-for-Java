package com.aspose.cells.examples.CellsHelperClass;

import com.aspose.cells.CellsHelper;
import com.aspose.cells.LoadDataOption;
import com.aspose.cells.LoadOptions;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class MergeFiles {

	public static void main(String[] args) throws Exception {

		String dataDir = Utils.getSharedDataDir(MergeFiles.class) + "CellsHelperClass/";
		String sampleFile = "Sample.out.xlsx";
		String samplePath = dataDir + sampleFile;

		// Create an Array (length=2)
		String[] files = new String[2];
		// Specify files with their paths to be merged
		files[0] = dataDir + "Book1.xls";
		files[1] = dataDir + "Book2.xls";

		// Create a cachedFile for the process
		String cacheFile = dataDir + "test.txt";
		// Output File to be created
		String dest = dataDir + "output.xls";

		// Merge the files in the output file
		CellsHelper.mergeFiles(files, cacheFile, dest);

		// Now if you need to rename your sheets, you may load the output file
		Workbook workbook = new Workbook(dataDir + "output.xls");

		int cnt = 1;

		// Browse all the sheets to rename them accordingly
		for (int i = 0; i < workbook.getWorksheets().getCount(); i++) {
			workbook.getWorksheets().get(i).setName("Sheet1" + cnt);
			cnt++;
		}
		// Re-save the file
		workbook.save(dataDir + "MergeFiles-out.xls");

	}
}
