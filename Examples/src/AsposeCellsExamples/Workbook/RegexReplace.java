package AsposeCellsExamples.Workbook;

import AsposeCellsExamples.Utils;
import com.aspose.cells.*;

public class RegexReplace {

	public static void main(String[] args) throws Exception {
        // ExStart: 1
        // directories
        String sourceDir = Utils.Get_SourceDirectory();
        String outputDir = Utils.Get_OutputDirectory();

		Workbook workbook = new Workbook(sourceDir + "SampleRegexReplace.xlsx");

        ReplaceOptions replace = new ReplaceOptions();
        replace.setCaseSensitive(false);
        replace.setMatchEntireCellContents(false);
        // Set to true to indicate that the searched key is regex
        replace.setRegexKey(true);

        workbook.replace("\\bKIM\\b", "^^^TIM^^^", replace);
        workbook.save(outputDir + "RegexReplace_out.xlsx");
        // ExEnd:1

		System.out.println("RegexReplace executed successfully.");
	}
}
