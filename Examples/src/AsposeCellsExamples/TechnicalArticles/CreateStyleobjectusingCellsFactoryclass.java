package AsposeCellsExamples.TechnicalArticles;

import com.aspose.cells.BackgroundType;
import com.aspose.cells.CellsFactory;
import com.aspose.cells.Color;
import com.aspose.cells.Style;
import com.aspose.cells.Workbook;
import AsposeCellsExamples.Utils;

public class CreateStyleobjectusingCellsFactoryclass {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(CreateStyleobjectusingCellsFactoryclass.class) + "articles/";

		// Create a Style object using CellsFactory class
		CellsFactory cf = new CellsFactory();
		Style st = cf.createStyle();

		// Set the Style fill color to Yellow
		st.setPattern(BackgroundType.SOLID);
		st.setForegroundColor(Color.getYellow());

		// Create a workbook and set its default style using the created Style
		// object
		Workbook wb = new Workbook();
		wb.setDefaultStyle(st);

		// Save the workbook
		wb.save(dataDir + "CreateStyleobject_out.xlsx");
	}
}
