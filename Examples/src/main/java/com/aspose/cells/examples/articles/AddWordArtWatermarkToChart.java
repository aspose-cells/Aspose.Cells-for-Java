package com.aspose.cells.examples.articles;

import com.aspose.cells.Chart;
import com.aspose.cells.FillFormat;
import com.aspose.cells.LineFormat;
import com.aspose.cells.MsoLineFormat;
import com.aspose.cells.MsoPresetTextEffect;
import com.aspose.cells.Shape;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class AddWordArtWatermarkToChart {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(AddWordArtWatermarkToChart.class) + "articles/";
		// Instantiate a new workbook, Open the existing excel file.

		Workbook workbook = new Workbook(dataDir + "sample.xlsx");

		// Get the chart in the first worksheet.
		Chart chart = workbook.getWorksheets().get(0).getCharts().get(0);

		// Add a WordArt watermark (shape) to the chart's plot area.
		Shape wordart = chart.getShapes().addTextEffectInChart(MsoPresetTextEffect.TEXT_EFFECT_1, "CONFIDENTIAL",
				"Arial Black", 66, false, false, 1200, 500, 2000, 3000);

		// Get the shape's fill format.
		FillFormat wordArtFormat = wordart.getFill();

		// Set the transparency.
		wordArtFormat.setTransparency(0.9);

		// Get the line format.
		MsoLineFormat lineFormat = wordart.getLineFormat();

		// Set Line format to invisible.
		lineFormat.setWeight(0.0);

		// Save the excel file.
		workbook.save(dataDir + "AWArtWToC_out.xlsx");

	}
}
