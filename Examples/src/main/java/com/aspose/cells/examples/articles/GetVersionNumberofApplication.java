package com.aspose.cells.examples.articles;

import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class GetVersionNumberofApplication {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(GetVersionNumberofApplication.class) + "articles/";
		// Create a workbook reference
		Workbook workbook = null;

		// Print the version number of Excel 2003 XLS file
		workbook = new Workbook(dataDir + "Excel2003.xls");
		System.out.println("Excel 2003 XLS Version: " + workbook.getBuiltInDocumentProperties().getVersion());

		// Print the version number of Excel 2007 XLS file
		workbook = new Workbook(dataDir + "Excel2007.xls");
		System.out.println("Excel 2007 XLS Version: " + workbook.getBuiltInDocumentProperties().getVersion());

		// Print the version number of Excel 2010 XLS file
		workbook = new Workbook(dataDir + "Excel2010.xls");
		System.out.println("Excel 2010 XLS Version: " + workbook.getBuiltInDocumentProperties().getVersion());

		// Print the version number of Excel 2013 XLS file
		workbook = new Workbook(dataDir + "Excel2013.xls");
		System.out.println("Excel 2013 XLS Version: " + workbook.getBuiltInDocumentProperties().getVersion());

		// Print the version number of Excel 2007 XLSX file
		workbook = new Workbook(dataDir + "Excel2007.xlsx");
		System.out.println("Excel 2007 XLSX Version: " + workbook.getBuiltInDocumentProperties().getVersion());

		// Print the version number of Excel 2010 XLSX file
		workbook = new Workbook(dataDir + "Excel2010.xlsx");
		System.out.println("Excel 2010 XLSX Version: " + workbook.getBuiltInDocumentProperties().getVersion());

		// Print the version number of Excel 2013 XLSX file
		workbook = new Workbook(dataDir + "Excel2013.xlsx");
		System.out.println("Excel 2013 XLSX Version: " + workbook.getBuiltInDocumentProperties().getVersion());

	}
}
