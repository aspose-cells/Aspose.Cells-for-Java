package AsposeCellsExamples.Workbook;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class FilterDefinedNamesWhileLoadingWorkbook {
	
	static String srcDir = Utils.Get_SourceDirectory();
	static String outDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {

		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());
		
		//Specify the load options
		LoadOptions opts = new LoadOptions();
		 
		//We do not want to load defined names
		opts.setLoadFilter(new LoadFilter(~LoadDataFilterOptions.DEFINED_NAMES));
		 
		//Load the workbook
		Workbook wb = new Workbook(srcDir + "sampleFilterDefinedNamesWhileLoadingWorkbook.xlsx", opts);
		 
		//Save the output Excel file, it will break the formula in C1
		wb.save(outDir + "outputFilterDefinedNamesWhileLoadingWorkbook.xlsx");

		// Print the message
		System.out.println("FilterDefinedNamesWhileLoadingWorkbook executed successfully.");
	}
}
