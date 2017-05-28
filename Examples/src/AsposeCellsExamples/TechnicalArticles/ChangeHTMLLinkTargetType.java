package AsposeCellsExamples.TechnicalArticles;

import com.aspose.cells.HtmlLinkTargetType;
import com.aspose.cells.HtmlSaveOptions;
import com.aspose.cells.Workbook;
import AsposeCellsExamples.Utils;

public class ChangeHTMLLinkTargetType {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ChangeHTMLLinkTargetType.class) + "TechnicalArticles/";
		Workbook workbook = new Workbook(dataDir + "Sample1.xlsx");

		HtmlSaveOptions opts = new HtmlSaveOptions();
		opts.setLinkTargetType(HtmlLinkTargetType.SELF);

		workbook.save(dataDir + "out.html", opts);

	}
}
