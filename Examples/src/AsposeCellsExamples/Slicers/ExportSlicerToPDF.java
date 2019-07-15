package AsposeCellsExamples.Slicers;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class ExportSlicerToPDF {
	
	static String sourceDir = Utils.Get_SourceDirectory();
	static String outputDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {

		// ExStart:1
        Workbook workbook = new Workbook(sourceDir + "SampleSlicerChart.xlsx");
        workbook.save(outputDir + "SampleSlicerChart.pdf", SaveFormat.PDF);
        // ExEnd:1
		
		System.out.println("ExportSlicerToPDF executed successfully.");
	}
}
