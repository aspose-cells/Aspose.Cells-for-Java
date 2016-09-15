package com.aspose.cells.examples.articles;

import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class CheckVBAProjectInWorkbookIsSigned {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(CheckVBAProjectInWorkbookIsSigned.class) + "articles/";
		Workbook workbook = new Workbook(dataDir + "source.xlsm");
		System.out.println("VBA Project is Signed: " + workbook.getVbaProject().isSigned());


	}
}
