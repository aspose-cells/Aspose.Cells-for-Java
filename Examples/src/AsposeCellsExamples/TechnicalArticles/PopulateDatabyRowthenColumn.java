package AsposeCellsExamples.TechnicalArticles;

import com.aspose.cells.Cells;
import com.aspose.cells.Workbook;

public class PopulateDatabyRowthenColumn {
	public static void main(String[] args) throws Exception {

		Workbook workbook = new Workbook();
		Cells cells = workbook.getWorksheets().get(0).getCells();
		cells.get("A1").setValue("data1");
		cells.get("B1").setValue("data2");
		cells.get("A2").setValue("data3");
		cells.get("B2").setValue("data4");

	}
}
