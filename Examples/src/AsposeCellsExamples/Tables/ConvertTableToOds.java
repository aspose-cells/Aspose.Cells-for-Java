package AsposeCellsExamples.Tables;

import AsposeCellsExamples.Utils;
import com.aspose.cells.CellsHelper;
import com.aspose.cells.Workbook;

public class ConvertTableToOds {
	public static void main(String[] args) throws Exception {
		System.out.println("Aspose.Cells for Java Version: " + CellsHelper.getVersion());
		// ExStart:1
		//Source directory
		String sourceDir = Utils.Get_SourceDirectory();

		//Output directory
		String outputDir = Utils.Get_OutputDirectory();

		// Open an existing file that contains a table/list object in it
		Workbook workbook = new Workbook(sourceDir + "SampleTable.xlsx");

		// Save the file
		workbook.save(outputDir + "ConvertTableToOds_out.ods");
		// ExEnd:1

		System.out.println("ConvertTableToOds executed successfully.");
	}
}
