package AsposeCellsExamples.Charts;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class SetValuesFormatCodeOfChartSeries {
	
	static String srcDir = Utils.Get_SourceDirectory();
	static String outDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {

		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());
			
		//Load the source Excel file 
		Workbook wb = new Workbook(srcDir + "sampleSeries_ValuesFormatCode.xlsx");
		  
		//Access first worksheet
		Worksheet worksheet = wb.getWorksheets().get(0);
		  
		//Access first chart
		Chart ch = worksheet.getCharts().get(0);
		  
		//Add series using an array of values
		ch.getNSeries().add("{10000, 20000, 30000, 40000}", true);
		  
		//Access the series and set its values format code
		Series srs = ch.getNSeries().get(0);
		srs.setValuesFormatCode("$#,##0");
		  
		//Save the output Excel file
		wb.save(outDir + "outputSeries_ValuesFormatCode.xlsx");

		// Print the message
		System.out.println("SetValuesFormatCodeOfChartSeries executed successfully.");
	}
}
