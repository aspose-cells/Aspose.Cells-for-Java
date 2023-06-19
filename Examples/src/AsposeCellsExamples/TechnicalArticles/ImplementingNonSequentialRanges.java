package AsposeCellsExamples.TechnicalArticles;

import com.aspose.cells.Name;
import com.aspose.cells.Workbook;
import AsposeCellsExamples.Utils;

public class ImplementingNonSequentialRanges {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ImplementingNonSequentialRanges.class) + "TechnicalArticles/";
		// Instantiating a Workbook object
		Workbook workbook = new Workbook();

		// Adding a Name for non sequenced range
		int index = workbook.getWorksheets().getNames().add("NonSequencedRange");

		Name name = workbook.getWorksheets().getNames().get(index);

		// Creating a non sequence range of cells
		name.setRefersTo("=Sheet1!$A$1:$B$3,Sheet1!$D$5:$E$6");

		// Save the workbook
		workbook.save(dataDir + "INSRanges_out.xls");

	}
}
