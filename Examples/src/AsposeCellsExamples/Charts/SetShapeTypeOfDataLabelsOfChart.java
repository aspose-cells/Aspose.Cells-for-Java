package AsposeCellsExamples.Charts;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class SetShapeTypeOfDataLabelsOfChart { 
	
	static String srcDir = Utils.Get_SourceDirectory();
	static String outDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {

		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());
		
		//Load source Excel file
		Workbook wb = new Workbook(srcDir + "sampleSetShapeTypeOfDataLabelsOfChart.xlsx");
		 
		//Access first worksheet
		Worksheet ws = wb.getWorksheets().get(0);
		 
		//Access first chart
		Chart ch = ws.getCharts().get(0);
		 
		//Access first series
		Series srs = ch.getNSeries().get(0);
		 
		//Set the shape type of data labels i.e. Speech Bubble Oval
		srs.getDataLabels().setShapeType(DataLabelShapeType.WEDGE_ELLIPSE_CALLOUT);
		 
		//Save the output Excel file
		wb.save(outDir + "outputSetShapeTypeOfDataLabelsOfChart.xlsx");

		// Print the message
		System.out.println("SetShapeTypeOfDataLabelsOfChart executed successfully.");
	}
}
