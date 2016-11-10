package com.aspose.cells.examples.articles;

import com.aspose.cells.ExternalLinkCollection;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class CheckWorkbookContainsHiddenExternalLinks {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(CheckWorkbookContainsHiddenExternalLinks.class) + "articles/";

		// Loads the workbook which contains hidden external links
		Workbook workbook = new Workbook(dataDir + "CheckWorkbookContainsHiddenExternalLinks_in.xlsx");

		// Access the external link collection of the workbook
		ExternalLinkCollection links = workbook.getWorksheets().getExternalLinks();

		// Print all the external links and check there IsVisible property
		for (int i = 0; i < links.getCount(); i++) {
			System.out.println("Data Source: " + links.get(i).getDataSource());
			System.out.println("Is Referred: " + links.get(i).isReferred());
			System.out.println("Is Visible: " + links.get(i).isVisible());
			System.out.println();
		}
	}
}
