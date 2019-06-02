package AsposeCellsExamples.Worksheets;

import com.aspose.cells.Color;
import com.aspose.cells.ODSPageBackground;
import com.aspose.cells.ODSPageBackgroundType;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import AsposeCellsExamples.Utils;
public class SetODSColoredBackground {
	
	public static void main(String[] args) throws Exception {
		// ExStart:1
		// The path to the output directory.
		String outDir = Utils.Get_OutputDirectory();
		
		// Instantiating a Workbook object
        Workbook workbook = new Workbook();

        //Access first worksheet
        Worksheet worksheet = workbook.getWorksheets().get(0);

        worksheet.getCells().get(0, 0).setValue(1);
        worksheet.getCells().get(1, 0).setValue(2);
        worksheet.getCells().get(2, 0).setValue(3);
        worksheet.getCells().get(3, 0).setValue(4);
        worksheet.getCells().get(4, 0).setValue(5);
        worksheet.getCells().get(5, 0).setValue(6);
        worksheet.getCells().get(0, 1).setValue(7);
        worksheet.getCells().get(1, 1).setValue(8);
        worksheet.getCells().get(2, 1).setValue(9);
        worksheet.getCells().get(3, 1).setValue(10);
        worksheet.getCells().get(4, 1).setValue(11);
        worksheet.getCells().get(5, 1).setValue(12);

        ODSPageBackground background = worksheet.getPageSetup().getODSPageBackground();

        background.setColor(Color.getAzure());
        background.setType(ODSPageBackgroundType.COLOR);

        workbook.save(outDir + "ColoredBackground.ods", SaveFormat.ODS);
        // ExEnd:1

        System.out.println("SetODSColoredBackground executed successfully.");
	}
}
