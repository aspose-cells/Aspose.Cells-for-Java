package AsposeCellsExamples.TechnicalArticles;

import AsposeCellsExamples.Utils;
import com.aspose.cells.*;

public class SettingScaleCropLinksUpToDate {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(SettingScaleCropLinksUpToDate.class) + "TechnicalArticles/";

		// Create workbook
		Workbook wb = new Workbook();

		// Setting ScaleCrop and LinksUpToDate BuiltInDocumentProperties
		wb.getBuiltInDocumentProperties().setScaleCrop(true);
		wb.getBuiltInDocumentProperties().setLinksUpToDate(true);

		// Save output excel file
		wb.save(dataDir + "output.xlsx");

	}
}