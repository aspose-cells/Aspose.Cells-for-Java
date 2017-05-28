package AsposeCellsExamples.Charts;

import java.io.FileInputStream;

import com.aspose.cells.Chart;
import com.aspose.cells.Color;
import com.aspose.cells.FillType;
import com.aspose.cells.LineFormat;
import com.aspose.cells.MsoLineDashStyle;
import com.aspose.cells.Picture;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import AsposeCellsExamples.Utils;

public class AddingPictureToChart {

	public static void main(String[] args) throws Exception {
		// ExStart:AddingPictureToChart
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(AddingPictureToChart.class) + "Charts/";

		String filePath = dataDir + "chart.xls";

		FileInputStream stream = new FileInputStream(dataDir + "logo.jpg");

		Workbook workbook = new Workbook(filePath);

		Worksheet worksheet = workbook.getWorksheets().get(0);

		// Load the chart from source worksheet
		Chart chart = worksheet.getCharts().get(0);

		Picture pic = chart.getShapes().addPictureInChart(50, 50, stream, 40, 40);
		LineFormat lineformat = pic.getLine();

		lineformat.setFillType(FillType.SOLID);
		lineformat.getSolidFill().setColor(Color.getBlue());
		lineformat.setDashStyle(MsoLineDashStyle.DASH_DOT_DOT);
		// Output the file
		workbook.save(dataDir + "APToChart_out.xls");

		// Print message
		System.out.println("Picture added to chart successfully.");
		// ExEnd:AddingPictureToChart
	}
}
