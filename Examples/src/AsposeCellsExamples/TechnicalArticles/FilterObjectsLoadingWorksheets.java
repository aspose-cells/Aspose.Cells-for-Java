package AsposeCellsExamples.TechnicalArticles;

import AsposeCellsExamples.Utils;
import com.aspose.cells.*;

// ExStart:1
public class FilterObjectsLoadingWorksheets {

	// Implement your own custom load filter, it will enable you to filter your
	// individual worksheet
	class CustomLoadFilter extends LoadFilter {
		public void startSheet(Worksheet sheet) {

			if (sheet.getName().equals("NoCharts")) {
				// Load everything and filter charts
				this.setLoadDataFilterOptions(LoadDataFilterOptions.ALL& ~LoadDataFilterOptions.CHART);
			}

			if (sheet.getName().equals("NoShapes")) {
				// Load everything and filter shapes
				this.setLoadDataFilterOptions(LoadDataFilterOptions.ALL& ~LoadDataFilterOptions.DRAWING);
			}

			if (sheet.getName().equals("NoConditionalFormatting")) {
				// Load everything and filter conditional formatting
				this.setLoadDataFilterOptions(LoadDataFilterOptions.ALL& ~LoadDataFilterOptions.CONDITIONAL_FORMATTING);
			}
		}// End StartSheet method.
	}// End CustomLoadFilter class.

	public static void main(String[] args) throws Exception {

		FilterObjectsLoadingWorksheets pg = new FilterObjectsLoadingWorksheets();
		pg.Run();
	}

	public void Run() throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(FilterObjectsLoadingWorksheets.class) + "TechnicalArticles/";

		// Filter worksheets using custom load filter
		LoadOptions ldOpts = new LoadOptions();
		ldOpts.setLoadFilter(new CustomLoadFilter());

		// Load the workbook with above filter
		Workbook wb = new Workbook(dataDir + "sampleFilterDifferentObjects.xlsx", ldOpts);

		// Take the image of all worksheets one by one
		for (int i = 0; i < wb.getWorksheets().getCount(); i++) {
			// Access worksheet at index i
			Worksheet ws = wb.getWorksheets().get(i);

			// Create image or print options, we want the image of entire
			// worksheet
			ImageOrPrintOptions opts = new ImageOrPrintOptions();
			opts.setOnePagePerSheet(true);
			opts.setImageType(ImageType.PNG);

			// Convert worksheet into image
			SheetRender sr = new SheetRender(ws, opts);
			sr.toImage(0, dataDir + ws.getName() + ".png");
		}
	}
}
// ExEnd:1