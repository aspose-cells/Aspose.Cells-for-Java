package asposefeatures.workingwithworksheets.convertspreadsheettopdf.java;

import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;

public class AsposeConverter
{
	public static void main(String[] args) throws Exception
	{
		String dataPath = "src/asposefeatures/workingwithworksheets/convertspreadsheettopdf/data/";

		Workbook workbook = new Workbook(dataPath + "workbook.xls");

		// Save the document in PDF format
		workbook.save(dataPath + "AsposeConvert.pdf", SaveFormat.PDF);

		// Print message
		System.out.println("Excel to PDF conversion performed successfully.");
	}
}