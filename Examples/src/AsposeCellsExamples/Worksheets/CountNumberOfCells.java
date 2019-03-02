package AsposeCellsExamples.Worksheets;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

import java.io.FileInputStream;

public class CountNumberOfCells {

	public static void main(String[] args) throws Exception {

		// ExStart:1
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(CountNumberOfCells.class) + "Worksheets/";
		//Load source Excel file
        Workbook workbook = new Workbook(dataDir + "BookWithSomeData.xlsx");

        //Access first worksheet
        Worksheet worksheet = workbook.getWorksheets().get(0);

        //Print number of cells in the Worksheet
        System.out.println("Number of Cells: " + worksheet.getCells().getCount());

        // If the number of cells is greater than 2147483647, use CountLarge
        System.out.println("Number of Cells (CountLarge): " + worksheet.getCells().getCountLarge());
        // ExEnd:1
	}
}
