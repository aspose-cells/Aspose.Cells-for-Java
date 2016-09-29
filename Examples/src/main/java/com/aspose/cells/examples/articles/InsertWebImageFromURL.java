package com.aspose.cells.examples.articles;

import java.io.BufferedInputStream;
import java.io.InputStream;
import java.net.URL;

import com.aspose.cells.PictureCollection;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class InsertWebImageFromURL {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(InsertWebImageFromURL.class) + "articles/";
		// Download image and store it in an object of InputStream
		URL url = new URL("https://www.google.com/images/nav_logo100633543.png");
		InputStream inStream = new BufferedInputStream(url.openStream());

		// Create a new workbook
		Workbook book = new Workbook();

		// Get the first worksheet in the book
		Worksheet sheet = book.getWorksheets().get(0);

		// Get the first worksheet pictures collection
		PictureCollection pictures = sheet.getPictures();

		// Insert the picture from the stream to B2 cell
		pictures.add(1, 1, inStream);

		// Save the excel file
		book.save(dataDir + "IWebImageFromURL_out.xls");

	}
}
