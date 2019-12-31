package AsposeCellsExamples.Workbook;

import AsposeCellsExamples.HelperClasses.DataObject;
import AsposeCellsExamples.Utils;
import com.aspose.cells.Workbook;
import com.aspose.cells.WorkbookDesigner;

import java.util.ArrayList;

public class AllowLeadingApostrophe {

	public static void main(String[] args) throws Exception {
		// ExStart:1
		//directories
		String sourceDir = Utils.Get_SourceDirectory();
		String outputDir = Utils.Get_OutputDirectory();

		// Instantiating a WorkbookDesigner object
		WorkbookDesigner designer = new WorkbookDesigner();

		Workbook workbook = new Workbook(sourceDir + "AllowLeadingApostropheSample.xlsx");
		workbook.getSettings().setQuotePrefixToStyle(false);

		// Open a designer spreadsheet containing smart markers
		designer.setWorkbook(workbook);

		ArrayList<DataObject> list = new ArrayList<>();
		list.add(new DataObject(1, "demo"));
		list.add(new DataObject(2, "'demo"));

		// Set the data source for the designer spreadsheet
		designer.setDataSource("sampleData", list);

		// Process the smart markers
		designer.process();

		designer.getWorkbook().save(outputDir + "AllowLeadingApostropheSample_out.xlsx");
		// ExEnd:1

		System.out.println("AllowLeadingApostrophe executed successfully.");
	}
}