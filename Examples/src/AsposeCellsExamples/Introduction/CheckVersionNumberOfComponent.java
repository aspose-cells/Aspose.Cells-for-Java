package AsposeCellsExamples.Introduction;

import com.aspose.cells.CellsHelper;
import com.aspose.cells.Workbook;

public class CheckVersionNumberOfComponent {
	public static void main(String[] args) throws Exception {
		try {
			System.out.println(CellsHelper.getVersion());
		}
		catch (Exception ee) {
			System.out.println(ee);
		}
	}
}