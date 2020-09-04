package AsposeCellsExamples.PivotTables;

import com.aspose.cells.PivotTable;
import com.aspose.cells.Workbook;
import AsposeCellsExamples.Utils;

public class ShowReportFilterPagesOption {
	static String srcDir = Utils.Get_SourceDirectory();
	static String outDir = Utils.Get_OutputDirectory();

	public static void main(String[] args) throws Exception {
        // ExStart:1
        // Load template file
        Workbook wb = new Workbook(srcDir + "samplePivotTable.xlsx");

        // Get first pivot table in the worksheet
        PivotTable pt = wb.getWorksheets().get(1).getPivotTables().get(0);

        // Set pivot field
        pt.showReportFilterPage(pt.getPageFields().get(0));

        // Set position index for showing report filter pages
        pt.showReportFilterPageByIndex(pt.getPageFields().get(0).getPosition());

        // Set the page field name
        pt.showReportFilterPageByName(pt.getPageFields().get(0).getName());

        // Save the output file
        wb.save(outDir + "outputSamplePivotTable.xlsx");
        // ExEnd:1
        
        System.out.println("ShowReportFilterPagesOption executed successfully.");
	}
}
