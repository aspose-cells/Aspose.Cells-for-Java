package com.aspose.cells.examples.articles;

import com.aspose.cells.ImageFormat;
import com.aspose.cells.ImageOrPrintOptions;
import com.aspose.cells.Picture;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class ExtractImagesfromWorksheets {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ExtractImagesfromWorksheets.class) + "articles/";
		// Open a template Excel file
		Workbook workbook = new Workbook(dataDir + "book1.xls");

		// Get the first worksheet
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Get the first Picture in the first worksheet
		Picture pic = worksheet.getPictures().get(0);

		// Set the output image file path
		String fileName = "aspose-logo.Jpg";
		String picformat = pic.getImageFormat().toString();

		// Note: you may evaluate the image format before specifying the image path

		// Define ImageOrPrintOptions
		ImageOrPrintOptions printoption = new ImageOrPrintOptions();

		// Specify the image format
		printoption.setImageFormat(ImageFormat.getJpeg());

		// Save the image
		pic.toImage(dataDir + fileName + picformat, printoption);

	}
}
