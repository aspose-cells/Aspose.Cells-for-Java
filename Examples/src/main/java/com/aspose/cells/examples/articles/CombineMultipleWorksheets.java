package com.aspose.cells.examples.articles;

import com.aspose.cells.Range;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class CombineMultipleWorksheets {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(CombineMultipleWorksheets.class) + "articles/";

		Workbook workbook = new Workbook(dataDir + "source.xlsx");

		Workbook destWorkbook = new Workbook();

		Worksheet destSheet = destWorkbook.getWorksheets().get(0);

		int TotalRowCount = 0;

		for (int i = 0; i < workbook.getWorksheets().getCount(); i++) {
			Worksheet sourceSheet = workbook.getWorksheets().get(i);

			Range sourceRange = sourceSheet.getCells().getMaxDisplayRange();

			Range destRange = destSheet.getCells().createRange(sourceRange.getFirstRow() + TotalRowCount,
					sourceRange.getFirstColumn(), sourceRange.getRowCount(), sourceRange.getColumnCount());

			destRange.copy(sourceRange);

			TotalRowCount = sourceRange.getRowCount() + TotalRowCount;
		}

		destWorkbook.save(dataDir + "CMWorksheets_out.xlsx");

	}
}
