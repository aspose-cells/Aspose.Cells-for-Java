package AsposeCellsExamples.Workbook;

import AsposeCellsExamples.Utils;
import com.aspose.cells.*;

public class AdjustCompressionLevel {

	public static void main(String[] args) throws Exception {
		// ExStart:1
		// The path to the directories.
		String sourceDir = Utils.Get_SourceDirectory();
		String outDir = Utils.Get_OutputDirectory();

		Workbook workbook = new Workbook(sourceDir + "LargeSampleFile.xlsx");
		XlsbSaveOptions options = new XlsbSaveOptions();
        options.setCompressionType(OoxmlCompressionType.LEVEL_1);
        long startTime = System.nanoTime();
        workbook.save(outDir + "LargeSampleFile_level_1_out.xlsb", options);
        long endTime = System.nanoTime();
        long timeElapsed = endTime - startTime;
        System.out.println("Level 1 Elapsed Time: " + timeElapsed / 1000000);

        startTime = System.nanoTime();
        options.setCompressionType(OoxmlCompressionType.LEVEL_6);
        workbook.save(outDir + "LargeSampleFile_level_6_out.xlsb", options);
        endTime = System.nanoTime();
        timeElapsed = endTime - startTime;
        System.out.println("Level 6 Elapsed Time: " + timeElapsed / 1000000);

        startTime = System.nanoTime();
        options.setCompressionType(OoxmlCompressionType.LEVEL_9);
        workbook.save(outDir + "LargeSampleFile_level_9_out.xlsb", options);
        endTime = System.nanoTime();
        timeElapsed = endTime - startTime;
        System.out.println("Level 9 Elapsed Time: " + timeElapsed / 1000000);
        // ExEnd:1

		System.out.println("AdjustCompressionLevel executed successfully.");
	}
}
