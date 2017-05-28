package AsposeCellsExamples.TechnicalArticles;

import com.aspose.cells.IWarningCallback;
import com.aspose.cells.PdfSaveOptions;
import com.aspose.cells.WarningInfo;
import com.aspose.cells.WarningType;
import com.aspose.cells.Workbook;
import AsposeCellsExamples.Utils;



public class WarningCallback implements IWarningCallback {

	@Override
	public void warning(WarningInfo info) {
		if (info.getWarningType() == WarningType.FONT_SUBSTITUTION) {
			System.out.println("WARNING INFO: " + info.getDescription());
		}
	}

	// ........
	// ........

	static void Run() throws Exception {

		String dataDir = Utils.getDataDir(WarningCallback.class);
		Workbook workbook = new Workbook(dataDir + "source.xlsx");

		PdfSaveOptions options = new PdfSaveOptions();
		options.setWarningCallback(new WarningCallback());

		workbook.save(dataDir + "WarningCallback_out.pdf", options);
	}
}
