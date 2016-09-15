package com.aspose.cells.examples.articles;

import java.util.ArrayList;

import com.aspose.cells.Cell;
import com.aspose.cells.PdfBookmarkEntry;
import com.aspose.cells.PdfSaveOptions;
import com.aspose.cells.Workbook;
import com.aspose.cells.WorksheetCollection;
import com.aspose.cells.examples.Utils;

public class AddPDFBookmarks {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(AddPDFBookmarks.class) + "articles/";
		// Instantiate a new workbook.
		Workbook workbook = new Workbook();

		// Get the worksheets in the workbook.
		WorksheetCollection worksheets = workbook.getWorksheets();

		// Add a sheet to the workbook.
		worksheets.add("1");

		// Add 2nd sheet to the workbook.
		worksheets.add("2");

		// Add the third sheet.
		worksheets.add("3");

		// Get cells in different worksheets.
		Cell cellInPage1 = worksheets.get(0).getCells().get("A1");
		Cell cellInPage2 = worksheets.get(1).getCells().get("A1");
		Cell cellInPage3 = worksheets.get(2).getCells().get("A1");

		// Add a value to the A1 cell in the first sheet.
		cellInPage1.setValue("a");

		// Add a value to the A1 cell in the second sheet.
		cellInPage2.setValue("b");

		// Add a value to the A1 cell in the third sheet.
		cellInPage3.setValue("c");

		// Create the PdfBookmark entry object.
		PdfBookmarkEntry pbeRoot = new PdfBookmarkEntry();

		// Set its text.
		pbeRoot.setText("root");

		// Set its destination source page.
		pbeRoot.setDestination(cellInPage1);

		// Set the bookmark collapsed.
		pbeRoot.setOpen(false);

		// Add a new PdfBookmark entry object.
		PdfBookmarkEntry subPbe1 = new PdfBookmarkEntry();

		// Set its text.
		subPbe1.setText("1");

		// Set its destination source page.
		subPbe1.setDestination(cellInPage2);

		// Add another PdfBookmark entry object.
		PdfBookmarkEntry subPbe2 = new PdfBookmarkEntry();

		// Set its text.
		subPbe2.setText("2");

		// Set its destination source page.
		subPbe2.setDestination(cellInPage3);

		// Create an array list.
		ArrayList subEntryList = new ArrayList();

		// Add the entry objects to it.
		subEntryList.add(subPbe1);
		subEntryList.add(subPbe2);
		pbeRoot.setSubEntry(subEntryList);

		// Set the PDF bookmarks.
		PdfSaveOptions options = new PdfSaveOptions();
		options.setBookmark(pbeRoot);

		// Save the PDF file.
		workbook.save(dataDir, options);

	}
}
