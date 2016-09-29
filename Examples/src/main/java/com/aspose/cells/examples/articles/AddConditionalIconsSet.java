package com.aspose.cells.examples.articles;

import java.io.ByteArrayInputStream;

import com.aspose.cells.Cells;
import com.aspose.cells.ConditionalFormattingIcon;
import com.aspose.cells.IconSetType;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class AddConditionalIconsSet {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(AddConditionalIconsSet.class) + "articles/";
		// Instantiate an instance of Workbook
		Workbook workbook = new Workbook();
		// Get the first worksheet (default worksheet) in the workbook
		Worksheet worksheet = workbook.getWorksheets().get(0);
		// Get the cells
		Cells cells = worksheet.getCells();
		// Set the columns widths (A, B and C)
		worksheet.getCells().setColumnWidth(0, 24);
		worksheet.getCells().setColumnWidth(1, 24);
		worksheet.getCells().setColumnWidth(2, 24);

		// Input date into the cells
		cells.get("A1").setValue("KPIs");
		cells.get("A2").setValue("Total Turnover (Sales at List)");
		cells.get("A3").setValue("Total Gross Margin %");
		cells.get("A4").setValue("Total Net Margin %");
		cells.get("B1").setValue("UA Contract Size Group 4");
		cells.get("B2").setValue(19551794);
		cells.get("B3").setValue(11.8070745566204);
		cells.get("B4").setValue(11.858589818569);
		cells.get("C1").setValue("UA Contract Size Group 3");
		cells.get("C2").setValue(8150131.66666667);
		cells.get("C3").setValue(10.3168384396244);
		cells.get("C4").setValue(11.3326931937091);

		// Get the conditional icon's image data
		byte[] imagedata = ConditionalFormattingIcon.getIconImageData(IconSetType.TRAFFIC_LIGHTS_31, 0);
		// Create a stream based on the image data
		ByteArrayInputStream stream = new ByteArrayInputStream(imagedata);
		// Add the picture to the cell based on the stream
		worksheet.getPictures().add(1, 1, stream);

		// Get the conditional icon's image data
		byte[] imagedata1 = ConditionalFormattingIcon.getIconImageData(IconSetType.ARROWS_3, 2);
		// Create a stream based on the image data
		ByteArrayInputStream stream1 = new ByteArrayInputStream(imagedata1);
		// Add the picture to the cell based on the stream
		worksheet.getPictures().add(1, 2, stream1);

		// Get the conditional icon's image data
		byte[] imagedata2 = ConditionalFormattingIcon.getIconImageData(IconSetType.SYMBOLS_3, 0);
		// Create a stream based on the image data
		ByteArrayInputStream stream2 = new ByteArrayInputStream(imagedata2);
		// Add the picture to the cell based on the stream
		worksheet.getPictures().add(2, 1, stream2);

		// Get the conditional icon's image data
		byte[] imagedata3 = ConditionalFormattingIcon.getIconImageData(IconSetType.STARS_3, 0);
		// Create a stream based on the image data
		ByteArrayInputStream stream3 = new ByteArrayInputStream(imagedata3);
		// Add the picture to the cell based on the stream
		worksheet.getPictures().add(2, 2, stream3);

		// Get the conditional icon's image data
		byte[] imagedata4 = ConditionalFormattingIcon.getIconImageData(IconSetType.BOXES_5, 1);
		// Create a stream based on the image data
		ByteArrayInputStream stream4 = new ByteArrayInputStream(imagedata4);
		// Add the picture to the cell based on the stream
		worksheet.getPictures().add(3, 1, stream4);

		// Get the conditional icon's image data
		byte[] imagedata5 = ConditionalFormattingIcon.getIconImageData(IconSetType.FLAGS_3, 1);
		// Create a stream based on the image data
		ByteArrayInputStream stream5 = new ByteArrayInputStream(imagedata5);
		// Add the picture to the cell based on the stream
		worksheet.getPictures().add(3, 2, stream5);

		// Save the Excel file
		workbook.save(dataDir + "ACIconsSet_out.xlsx");

	}
}
