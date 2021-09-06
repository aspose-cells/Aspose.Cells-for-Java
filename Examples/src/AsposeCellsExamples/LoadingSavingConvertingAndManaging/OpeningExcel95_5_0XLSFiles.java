package AsposeCellsExamples.LoadingSavingConvertingAndManaging;

import com.aspose.cells.Workbook;

import AsposeCellsExamples.Utils;

public class OpeningExcel95_5_0XLSFiles {
	static String srcDir = Utils.Get_SourceDirectory();
	
	public static void main(String[] args) throws Exception {
		// ExStart:1
		// Opening Microsoft Excel 97 Files

		// Creating an Workbook object with excel 97 file path
		new Workbook(srcDir + "Excel95_5.0.xls");

		// ExEnd:1
		// Print message
		System.out.println("Excel 95/5.0 XLS Workbook opened successfully.");
	}
}
