package com.aspose.cells.examples.articles;

import com.aspose.cells.Chart;
import com.aspose.cells.Color;
import com.aspose.cells.DataLabels;
import com.aspose.cells.FontSetting;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class RichTextCustomData {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(RichTextCustomData.class) + "articles/";

		// Create a workbook from source Excel file
		Workbook workbook = new Workbook(dataDir + "source.xlsx");

		// Access first worksheet
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Access the first chart inside the sheet
		Chart chart = worksheet.getCharts().get(0);

		// Access the data label of first series first point
		DataLabels dlbls = chart.getNSeries().get(0).getPoints().get(0).getDataLabels();

		// Set data label text
		dlbls.setText("Rich Text Label");

		// Set the font setting of the first 10 characters
		FontSetting fntSetting = dlbls.characters(0, 10);
		fntSetting.getFont().setColor(Color.getRed());
		fntSetting.getFont().setBold(true);

		// Save the workbook
		workbook.save("RTCustomData_out.xlsx");

	}
}
