package AsposeCellsExamples.Charts;

import java.util.*;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class ReadManipulateExcel2016Charts {

	public static void main(String[] args) throws Exception {
	
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ReadManipulateExcel2016Charts.class) + "Charts/";

		//ExStart: 1
		// Load source excel file containing excel 2016 charts
		Workbook wb = new Workbook(dataDir + "excel2016Charts.xlsx");

		// Access the first worksheet which contains the charts
		Worksheet ws = wb.getWorksheets().get(0);

		//Converting integer enums to string enums
		HashMap<Integer, String> cTypes = new HashMap<>();
		cTypes.put(ChartType.BOX_WHISKER, "BoxWhisker");
		cTypes.put(ChartType.WATERFALL, "Waterfall");
		cTypes.put(ChartType.TREEMAP, "Treemap");
		cTypes.put(ChartType.HISTOGRAM, "Histogram");
		cTypes.put(ChartType.SUNBURST, "Sunburst");

		// Access all charts one by one and read their types
		for (int i = 0; i < ws.getCharts().getCount(); i++) {
			// Access the chart
			Chart ch = ws.getCharts().get(i);

			// Print chart type
			String strChartType = cTypes.get(ch.getType());
			System.out.println(strChartType);

			// Change the title of the charts as per their types
			ch.getTitle().setText("Chart Type is " + strChartType);
		}

		// Save the workbook
		wb.save(dataDir + "out_excel2016Charts.xlsx");
		// ExEnd: 1

		// Print message
		System.out.println("Excel 2016 Chart Titles changed successfully.");

	}
}
