package com.aspose.cells.examples.articles;

import java.util.ArrayList;

import com.aspose.cells.FontSetting;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Shape;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class ChangeCharacterSpacing {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ChangeCharacterSpacing.class) + "articles/";
		// Load your excel file inside a workbook obect
		Workbook wb = new Workbook(dataDir + "character-spacing.xlsx");

		// Access your text box which is also a shape object from shapes collection
		Shape shape = wb.getWorksheets().get(0).getShapes().get(0);

		// Access the first font setting object via GetCharacters() method
		ArrayList<FontSetting> lst = shape.getCharacters();
		FontSetting fs = lst.get(0);

		// Set the character spacing to point 4
		fs.getShapeFont().setSpacing(4);

		// Save the workbook in xlsx format
		wb.save(dataDir + "CCSpacing_out.xlsx", SaveFormat.XLSX);

	}

}
