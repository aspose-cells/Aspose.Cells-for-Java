package com.aspose.cells.examples.articles;

import com.aspose.cells.LoadOptions;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class CheckPassword {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(CheckPassword.class) + "articles/";

		// Specify password to open inside the load options
		LoadOptions opts = new LoadOptions();
		opts.setPassword("1234");

		// Open the source Excel file with load options
		Workbook workbook = new Workbook(dataDir + "Book1.xlsx", opts);

		// Check if 567 is Password to modify
		boolean ret = workbook.checkWriteProtectedPassword("567");
		System.out.println("Is 567 correct Password to modify: " + ret);

		// Check if 5679 is Password to modify
		ret = workbook.checkWriteProtectedPassword("5678");
		System.out.println("Is 5678 correct Password to modify: " + ret);

	}
}
