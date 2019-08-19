package AsposeCellsExamples.LoadingSavingConvertingAndManaging;

import java.io.ByteArrayInputStream;

import com.aspose.cells.HtmlLoadOptions;
import com.aspose.cells.Workbook;
import AsposeCellsExamples.Utils;

public class AutoFitColumnsRowsLoadingHTML {

	public static void main(String[] args) throws Exception 
	{
		// ExStart:1
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(AutoFitColumnsRowsLoadingHTML.class) + "LoadingSavingConvertingAndManaging/";
		
		//Sample HTML.
		String sampleHtml = "<html><body><table><tr><td>This is sample text.</td><td>Some text.</td></tr><tr><td>This is another sample text.</td><td>Some text.</td></tr></table></body></html>";
		//Load html string into byte array input stream
		ByteArrayInputStream bais = new ByteArrayInputStream(sampleHtml.getBytes());
		  
		//Load byte array stream into workbook.
		Workbook wb = new Workbook(bais);
		  
		//Save the workbook in xlsx format.
		wb.save(dataDir + "outputWithout_AutoFitColsAndRows.xlsx");
		  
		//Specify the HtmlLoadOptions and set AutoFitColsAndRows = true.
		HtmlLoadOptions opts = new HtmlLoadOptions();
		opts.setAutoFitColsAndRows(true);
		  
		//Load byte array stream into workbook with the above HtmlLoadOptions.
		bais.reset();
		wb = new Workbook(bais, opts);
		  
		//Save the workbook in xlsx format.
		wb.save(dataDir + "outputWith_AutoFitColsAndRows.xlsx");
		// ExEnd:1
	}

}
