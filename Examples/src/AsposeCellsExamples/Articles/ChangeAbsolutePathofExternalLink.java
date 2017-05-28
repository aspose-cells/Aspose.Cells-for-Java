package com.aspose.cells.examples.articles;

import com.aspose.cells.ExternalLink;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class ChangeAbsolutePathofExternalLink {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ChangeAbsolutePathofExternalLink.class) + "articles/";

		// Load your source excel file containing the external link
		Workbook wb = new Workbook(dataDir + "sample.xlsx");

		// Access the first external link
		ExternalLink externalLink = wb.getWorksheets().getExternalLinks().get(0);

		// Print the data source of external link, it will print existing remote
		// path
		System.out.println("External Link Data Source: " + externalLink.getDataSource());

		// Remove the remote path and print the new data source
		// Assign the new data source to external link and print again, it will
		// now print data source with local path
		externalLink.setDataSource("ExternalAccounts.xlsx");
		System.out.println("External Link Data Source After Removing Remote Path: " + externalLink.getDataSource());

		// Change the absolute path of the workbook, it will also change the
		// external link path
		wb.setAbsolutePath("C:\\Files\\Extra\\");

		// Now print the data source again
		System.out.println("External Link Data Source After Changing Workbook.AbsolutePath to Local Path: " + externalLink.getDataSource());

		// Change the absolute path of the workbook to some remote path, it will
		// again affect the external link path
		wb.setAbsolutePath("http://www.aspose.com/WebFiles/ExcelFiles/");

		// Now print the data source again
		System.out.println("External Link Data Source After Changing Workbook.AbsolutePath to Remote Path: " + externalLink.getDataSource());
	}
}
