package AsposeCellsExamples.LoadingSavingConvertingAndManaging;

import com.aspose.cells.LoadOptions;
import com.aspose.cells.Workbook;
import AsposeCellsExamples.Utils;

public class GetWarningLoadingAnExcel {
	public static void main(String[] args) throws Exception 
	{
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(GetWarningLoadingAnExcel.class) + "LoadingSavingConvertingAndManaging/";
		
		//Create load options and set the WarningCallback property 
		//to catch warnings while loading workbook
		LoadOptions options = new LoadOptions();
		options.setWarningCallback(new WarningCallback());
		              
		//Load the source excel file
		Workbook book = new Workbook(dataDir + "sampleDuplicateDefinedName.xlsx", options);
		  
		//Save the workbook 
		book.save(dataDir + "outputDuplicateDefinedName.xlsx");
		
	}
}
