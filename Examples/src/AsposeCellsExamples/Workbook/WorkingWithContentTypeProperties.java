package AsposeCellsExamples.Workbook;

import AsposeCellsExamples.Utils;
import com.aspose.cells.*;

public class WorkingWithContentTypeProperties {

	public static void main(String[] args) throws Exception {

        // ExStart:1
        // The path to the directories.
        String outputDir = Utils.Get_OutputDirectory();

        Workbook workbook = new Workbook(FileFormatType.XLSX);
        int index = workbook.getContentTypeProperties().add("MK31", "Simple Data");
        workbook.getContentTypeProperties().get(index).setNillable(true);
        index= workbook.getContentTypeProperties().add("MK32", "2019-10-17T16:00:00+00:00", "DateTime");
        workbook.getContentTypeProperties().get(index).setNillable(false);
        workbook.save(outputDir + "WorkingWithContentTypeProperties_out.xlsx");
        // ExEnd:1

		System.out.println("WorkingWithContentTypeProperties executed successfully.");
	}
}
