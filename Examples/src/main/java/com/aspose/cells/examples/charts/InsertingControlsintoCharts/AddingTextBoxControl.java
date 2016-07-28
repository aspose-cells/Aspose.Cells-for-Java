package com.aspose.cells.examples.charts.InsertingControlsintoCharts;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;

import java.io.FileInputStream;

public class AddingTextBoxControl {

	public static void main(String[] args) throws Exception {
		// ExStart:1
		// The path to the documents directory.
		String dataDir = Utils.getDataDir(AddingTextBoxControl.class);
		String filePath = dataDir + "chart.xls";

		// Create a new Workbook.
		// Open the existing file.
		Workbook workbook = new Workbook(filePath);
		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Load the chart from source worksheet
		Chart chart = worksheet.getCharts().get(0);

		// Add a new textbox to the chart.
		TextBox txt = chart.getShapes().addTextBoxInChart(100, 100, 850, 2500);
		txt.setText("Aspose");
		MsoTextFrame frame = txt.getTextFrame();
		frame.setAutoSize(true);
		txt.getFont().setItalic(true);
		txt.getFont().setSize(20);
		txt.getFont().setBold(true);

		// Get the filformat of the textbox.
		MsoFillFormat Fillformat = txt.getFillFormat();
		Fillformat.setForeColor(Color.getChocolate());

		// Get the lineformat type of the textbox.
		MsoLineFormat LineFormat = txt.getLineFormat();
		LineFormat.setWeight(2);
		LineFormat.setDashStyle(MsoLineDashStyle.SOLID);

		// Output the file
		workbook.save(dataDir + "output.xls");

		// Print message
		System.out.println("TextBox added to chart successfully.");
		// ExEnd:1
	}
}
