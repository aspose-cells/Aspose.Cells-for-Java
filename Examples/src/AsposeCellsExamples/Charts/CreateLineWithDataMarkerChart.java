package AsposeCellsExamples.Charts;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class CreateLineWithDataMarkerChart {

	static String outDir = Utils.Get_OutputDirectory();
	public static void main(String[] args) throws Exception {
        // Instantiate a workbook
        Workbook workbook = new Workbook();

        // Access first worksheet
        Worksheet worksheet = workbook.getWorksheets().get(0);

        // Set columns title 
        worksheet.getCells().get(0, 0).setValue("X");
        worksheet.getCells().get(0, 1).setValue("Y");

        // Create random data and save in the cells
        for (int i = 1; i < 21; i++)
        {
        	worksheet.getCells().get(i, 0).setValue(i);
        	worksheet.getCells().get(i, 1).setValue(0.8);
        }

        for (int i = 21; i < 41; i++)
        {
        	worksheet.getCells().get(i, 0).setValue(i - 20);
        	worksheet.getCells().get(i, 1).setValue(0.9);
        }
        // Add a chart to the worksheet
        int idx = worksheet.getCharts().add(ChartType.LINE_WITH_DATA_MARKERS, 1, 3, 20, 20);

        // Access the newly created chart
        Chart chart = worksheet.getCharts().get(idx);

        // Set chart style
        chart.setStyle(3);

        // Set autoscaling value to true
        chart.setAutoScaling(true);

        // Set foreground color white
        chart.getPlotArea().getArea().setForegroundColor(Color.getWhite());

        // Set Properties of chart title
        chart.getTitle().setText("Sample Chart");

        // Set chart type
        chart.setType(ChartType.LINE_WITH_DATA_MARKERS);

        // Set Properties of categoryaxis title
        chart.getCategoryAxis().getTitle().setText("Units");

        //Set Properties of nseries
        int s2_idx = chart.getNSeries().add("A2: A2", true);
        int s3_idx = chart.getNSeries().add("A22: A22", true);

        // Set IsColorVaried to true for varied points color
        chart.getNSeries().setColorVaried(true);

        // Set properties of background area and series markers
        chart.getNSeries().get(s2_idx).getArea().setFormatting(FormattingType.CUSTOM);
        chart.getNSeries().get(s2_idx).getMarker().getArea().setForegroundColor(Color.getYellow());
        chart.getNSeries().get(s2_idx).getMarker().getBorder().setVisible(false);

        // Set X and Y values of series chart
        chart.getNSeries().get(s2_idx).setXValues("A2: A21");
        chart.getNSeries().get(s2_idx).setValues("B2: B21");

        // Set properties of background area and series markers
        chart.getNSeries().get(s3_idx).getArea().setFormatting(FormattingType.CUSTOM);
        chart.getNSeries().get(s3_idx).getMarker().getArea().setForegroundColor(Color.getGreen());
        chart.getNSeries().get(s3_idx).getMarker().getBorder().setVisible(false);

        // Set X and Y values of series chart
        chart.getNSeries().get(s3_idx).setXValues("A22: A41");
        chart.getNSeries().get(s3_idx).setValues("B22: B41");

        // Save the workbook
        workbook.save(outDir + "LineWithDataMarkerChart.xlsx", SaveFormat.XLSX);

		// Print message
		System.out.println("Workbook with chart is successfully created.");
	}
}
