package AsposeCellsExamples.TechnicalArticles;

import com.aspose.cells.Workbook;
import AsposeCellsExamples.Utils;

public class GetFontsUsedinWorkbook {

	public static void main(String[] args)throws Exception 
	{
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(GetFontsUsedinWorkbook.class) + "articles/";
		//Load source workbook
		Workbook wb = new Workbook(dataDir + "sampleGetFonts.xlsx");
		  
		//Get all the fonts inside the workbook
		com.aspose.cells.Font[] fnts = wb.getFonts();
		  
		//Print all the fonts
		for(int i=0; i<fnts.length; i++)
		{
		    System.out.println(fnts[i]);
		}

	}

}
