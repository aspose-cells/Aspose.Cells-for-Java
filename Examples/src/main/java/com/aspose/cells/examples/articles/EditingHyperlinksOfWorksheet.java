package com.aspose.cells.examples.articles;

import com.aspose.cells.Hyperlink;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class EditingHyperlinksOfWorksheet {
	public static void main(String[] args) throws Exception {
		// ExStart:EditingHyperlinksOfWorksheet
		// The path to the documents directory.
		String dataDir = Utils.getDataDir(EditingHyperlinksOfWorksheet.class);
		Workbook workbook = new Workbook(dataDir + "source.xlsx");
		Worksheet worksheet = workbook.getWorksheets().get(0);
		for (int i = 0; i < worksheet.getHyperlinks().getCount(); i++) {
			Hyperlink hl = worksheet.getHyperlinks().get(i);
			hl.setAddress("http://www.aspose.com");
		}
		workbook.save(dataDir + "output.xlsx");
		// ExEnd:EditingHyperlinksOfWorksheet
	}
}
