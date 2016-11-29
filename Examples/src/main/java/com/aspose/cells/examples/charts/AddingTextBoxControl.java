package com.aspose.cells.examples.charts;

import com.aspose.cells.*;
import com.aspose.cells.examples.Utils;

public class AddingTextBoxControl {

	public static void main(String[] args) throws Exception {
		// ExStart:AddingTextBoxControl
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(AddingTextBoxControl.class) + "charts/";
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
		txt.getFont().setItalic(true);
		txt.getFont().setSize(20);
		txt.getFont().setBold(true);

		// Get the filformat of the textbox.
		FillFormat fillformat = txt.getFill();
		fillformat.setFillType(FillType.SOLID);
		fillformat.getSolidFill().setColor(Color.getSilver());

		// Get the lineformat type of the textbox.
		LineFormat lineformat = txt.getLine();
		lineformat.setWeight(2);
		lineformat.setDashStyle(MsoLineDashStyle.SOLID);

		// Output the file
		workbook.save(dataDir + "ATBoxControl_out.xls");

		// Print message
		System.out.println("TextBox added to chart successfully.");
		// ExEnd:AddingTextBoxControl
	}
}
