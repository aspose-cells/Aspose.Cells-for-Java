package com.aspose.cells.examples.articles;

import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class DeterminetheLicenseLoadedSuccessfully {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(DeterminetheLicenseLoadedSuccessfully.class) + "articles/";
		// Create workbook object before setting a license
		Workbook workbook = new Workbook();

		// Check if the license is loaded or not
		// It will return false
		System.out.println(workbook.isLicensed());

		// Set the license now
		String licPath = dataDir + "Aspose.Total.lic";

		com.aspose.cells.License lic = new com.aspose.cells.License();
		lic.setLicense(licPath);

		// Check if the license is loaded or not
		// Now it will return true
		System.out.println(workbook.isLicensed());

	}
}
