package AsposeCellsExamples.Charts;

import java.util.*;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class ReadAxisLabelsAfterCalculatingTheChart { 
	
	static String srcDir = Utils.Get_SourceDirectory();

	public static void main(String[] args) throws Exception {

		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());
		
		//Load the Excel file containing chart
		Workbook wb = new Workbook(srcDir + "sampleReadAxisLabelsAfterCalculatingTheChart.xlsx");

		//Access first worksheet
		Worksheet ws = wb.getWorksheets().get(0);

		//Access the chart
		Chart ch = ws.getCharts().get(0);

		//Calculate the chart
		ch.calculate();

		//Read axis labels of category axis
		ArrayList lstLabels = ch.getCategoryAxis().getAxisLabels();

		//Print axis labels on console
		System.out.println("Category Axis Labels: ");
		System.out.println("---------------------");

		//Iterate axis labels and print them one by one
		for(int i=0; i<lstLabels.size(); i++)
		{
			System.out.println(lstLabels.get(i));
		}

		// Print the message
		System.out.println("ReadAxisLabelsAfterCalculatingTheChart executed successfully.");
	}
}
