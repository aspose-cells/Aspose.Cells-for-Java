package com.aspose.cells.examples.articles;

import java.io.ByteArrayInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;

import com.aspose.cells.FileFormatType;
import com.aspose.cells.OleObject;
import com.aspose.cells.OleObjectCollection;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class ExtractOLEObjects {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ExtractOLEObjects.class) + "articles/";
		// Instantiating a Workbook object, Open the template file.
		Workbook workbook = new Workbook(dataDir + "oleFile.xlsx");

		// Get the OleObject Collection in the first worksheet.
		OleObjectCollection objects = workbook.getWorksheets().get(0).getOleObjects();

		// Loop through all the OleObjects and extract each object in the worksheet.
		for (int i = 0; i < objects.getCount(); i++) {
			OleObject object = objects.get(i);
			// Specify the output filename.
			String fileName = "D:/object" + i + ".";
			// Specify each file format based on the OleObject format type.

			switch (object.getFileFormatType()) {
			case FileFormatType.DOCX:
				fileName += "docx";
				break;
			case FileFormatType.XLSX:
				fileName += "xlsx";
				break;
			case FileFormatType.PPTX:
				fileName += "pptx";
				break;
			case FileFormatType.PDF:
				fileName += "pdf";
				break;
			case FileFormatType.UNKNOWN:
				fileName += "jpg";
				break;
			default:
				// ........
				break;
			}
			// Save the OleObject as a new excel file if the object type is xls.
			if (object.getFileFormatType() == FileFormatType.XLSX) {
				byte[] bytes = object.getObjectData();
				InputStream is = new ByteArrayInputStream(bytes);
				Workbook oleBook = new Workbook(is);
				oleBook.getSettings().setHidden(false);
				oleBook.save(fileName);
			}

			// Create the files based on the OleObject format types.
			else {
				FileOutputStream fos = new FileOutputStream(fileName);
				fos.write(object.getObjectData());
				fos.close();
			}
		}

	}
}
