package AsposeCellsExamples.Data;

import java.util.*;

import AsposeCellsExamples.HelperClasses.*;

import com.aspose.cells.*;
import AsposeCellsExamples.Utils;

public class ImportCustomObjectsToMergedArea {

	public static void main(String[] args) throws Exception {

		//ExStart: 1
		// The path to the source directory.
		String sourceDir = Utils.Get_SourceDirectory();
		// The path to the output directory.
		String outDir = Utils.Get_OutputDirectory();

		Workbook workbook = new Workbook(sourceDir + "sampleMergedTemplate.xlsx");
		Worksheet worksheet = workbook.getWorksheets().get(0);
		
		ArrayList productList = new ArrayList();
		
		for(int i = 0; i < 3; i++) {
			productList.add(new Product("Test Product - " + i, i*2));
		}
        
        ImportTableOptions tableOptions = new ImportTableOptions();
        tableOptions.setCheckMergedCells(true);
        tableOptions.setFieldNameShown(false);

        //Insert data to excel template
        worksheet.getCells().importCustomObjects(productList, 1, 0, tableOptions);
        workbook.save(outDir + "sampleMergedTemplate_out.xlsx", SaveFormat.XLSX);
        // ExEnd: 1

        System.out.println("ImportCustomObjectsToMergedArea executed successfully.");
	}
}
