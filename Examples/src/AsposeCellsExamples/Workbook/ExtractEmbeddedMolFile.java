package AsposeCellsExamples.Workbook;

import AsposeCellsExamples.Utils;
import com.aspose.cells.*;

import java.io.File;
import java.io.FileOutputStream;

public class ExtractEmbeddedMolFile {

	public static void main(String[] args) throws Exception {
		// ExStart:1
		// The path to the directories.
		String sourceDir = Utils.Get_SourceDirectory();
		String outputDir = Utils.Get_OutputDirectory();

        Workbook workbook = new Workbook(sourceDir + "EmbeddedMolSample.xlsx");
        int index = 1;
        for (Object obj : workbook.getWorksheets())
        {
            Worksheet sheet = (Worksheet)obj;
            OleObjectCollection oles = sheet.getOleObjects();
            for (Object obj2 : oles)
            {
                OleObject ole = (OleObject)obj2;
                String fileName = outputDir + "OleObject" + index + ".mol ";
                FileOutputStream fos = new FileOutputStream(fileName);
                fos.write(ole.getObjectData());
                fos.flush();
                fos.close();
                index++;
            }
        }
        // ExEnd:1

		System.out.println("ExtractEmbeddedMolFile executed successfully.");
	}
}
