package AsposeCellsExamples.WorkbookVBAProject;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class FilterVBAMacrosWhileLoadingWorkbook 
{
    static String srcDir = Utils.Get_SourceDirectory();
    static String outDir = Utils.Get_OutputDirectory();

    public static void main(String[] args) throws Exception
    {

        // Set the load options, we do not want to load VBA
        LoadOptions loadOptions = new LoadOptions(LoadFormat.AUTO);
        loadOptions.setLoadFilter(new LoadFilter(LoadDataFilterOptions.ALL & ~LoadDataFilterOptions.VBA));

        // Create workbook object from sample excel file using load options
        Workbook book = new Workbook(srcDir + "sampleMacroEnabledWorkbook.xlsm", loadOptions);

        // Save the output in pdf format
        book.save(outDir + "OutputSampleMacroEnabledWorkbook.xlsm", SaveFormat.XLSM);

        // Print the message
		System.out.println("FilterVBAMacrosWhileLoadingWorkbook executed successfully.");
	}
}
