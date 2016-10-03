package com.aspose.cells.examples.DrawingObjects;

import java.io.FileOutputStream;

import com.aspose.cells.FileFormatType;
import com.aspose.cells.MsoDrawingType;
import com.aspose.cells.OleObject;
import com.aspose.cells.OleObjectCollection;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class ExtractingOLEObjects {
	public static void main(String[] args) throws Exception {
		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(ExtractingOLEObjects.class) + "DrawingObjects/";

		// Instantiating a Workbook object
		Workbook workbook = new Workbook(dataDir + "book1.xls");

		// Get the OleObject Collection in the first worksheet.
		OleObjectCollection oles = workbook.getWorksheets().get(0).getOleObjects();

		// Loop through all the ole objects and extract each object. in the worksheet.
		for (int i = 0; i < oles.getCount(); i++) {
			if (oles.get(i).getMsoDrawingType() == MsoDrawingType.OLE_OBJECT) {
				OleObject ole = (OleObject) oles.get(i);
				// Specify the output filename.
				String fileName = dataDir + "tempBook1ole" + i + ".";
				// Specify each file format based on the oleformattype.
				switch (ole.getFileFormatType()) {
				case FileFormatType.DOC:
					fileName += "doc";
					break;
				case FileFormatType.EXCEL_97_TO_2003:
					fileName += "Xls";
					break;
				case FileFormatType.PPT:
					fileName += "Ppt";
					break;
				case FileFormatType.PDF:
					fileName += "Pdf";
					break;
				case FileFormatType.UNKNOWN:
					fileName += "Jpg";
					break;
				default:
					fileName += "data";
					break;
				}

				FileOutputStream fos = new FileOutputStream(fileName);
				byte[] data = ole.getObjectData();
				fos.write(data);
				fos.close();
			}
		}
	}
}
