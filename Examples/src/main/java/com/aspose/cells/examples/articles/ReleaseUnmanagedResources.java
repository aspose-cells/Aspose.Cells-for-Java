package com.aspose.cells.examples.articles;

import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class ReleaseUnmanagedResources {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ReleaseUnmanagedResources.class) + "articles/";
		// Create workbook object
		Workbook wb1 = new Workbook();

		/*
		 * Call dispose method,It performs application-defined tasks associated with freeing, releasing, or resetting
		 * unmanaged resources.
		 */
		wb1.dispose();


	}
}
