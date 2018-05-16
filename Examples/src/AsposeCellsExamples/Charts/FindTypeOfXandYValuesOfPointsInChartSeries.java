package AsposeCellsExamples.Charts;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class FindTypeOfXandYValuesOfPointsInChartSeries {
	
	static String srcDir = Utils.Get_SourceDirectory();
	static String outDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {

		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());
		
		//Converting integer enums to string enums
		java.util.HashMap<Integer, String> cvTypes = new java.util.HashMap<Integer, String>();
		cvTypes.put(CellValueType.IS_NUMERIC, "IsNumeric");
		cvTypes.put(CellValueType.IS_STRING, "IsString");

		//Load sample Excel file containing chart.
		Workbook wb = new Workbook(srcDir + "sampleFindTypeOfXandYValuesOfPointsInChartSeries.xlsx");

		//Access first worksheet.
		Worksheet ws = wb.getWorksheets().get(0);

		//Access first chart.
		Chart ch = ws.getCharts().get(0);

		//Calculate chart data.
		ch.calculate();

		//Access first chart point in the first series.
		ChartPoint pnt = ch.getNSeries().get(0).getPoints().get(0);

		//Print the types of X and Y values of chart point.
		System.out.println("X Value Type: " + cvTypes.get(pnt.getXValueType()));
		System.out.println("Y Value Type: " + cvTypes.get(pnt.getYValueType()));
		
		// Print the message
		System.out.println("FindTypeOfXandYValuesOfPointsInChartSeries executed successfully.");
	}
}
