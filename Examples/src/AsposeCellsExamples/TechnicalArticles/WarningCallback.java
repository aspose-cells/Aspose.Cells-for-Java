package AsposeCellsExamples.TechnicalArticles;

import com.aspose.cells.IWarningCallback;
import com.aspose.cells.PdfSaveOptions;
import com.aspose.cells.WarningInfo;
import com.aspose.cells.WarningType;
import com.aspose.cells.Workbook;
import AsposeCellsExamples.Utils;

public class WarningCallback implements IWarningCallback {
	// ExStart:1
	@Override
	public void warning(WarningInfo info) {
		if (info.getWarningType() == WarningType.FONT_SUBSTITUTION) {
			System.out.println("WARNING INFO: " + info.getDescription());
		}
	}

	// ........
	// ........

	public static void main(String[] args) throws Exception {

		String dataDir = Utils.getSharedDataDir(UseErrorCheckingOptions.class) + "TechnicalArticles/";
		Workbook workbook = new Workbook(dataDir + "source.xlsx");

		PdfSaveOptions options = new PdfSaveOptions();
		options.setWarningCallback(new WarningCallback());

		workbook.save(dataDir + "WarningCallback_out.pdf", options);
	}
	// ExEnd:1
}
