package AsposeCellsExamples.TechnicalArticles;

import com.aspose.cells.FileFormatInfo;
import com.aspose.cells.FileFormatType;
import com.aspose.cells.FileFormatUtil;

public class DetectFileFormatOfEncryptedFiles {
	public static void main(String[] args) throws Exception {

		// ExStart:1
		//Source directory
		String dataDir = AsposeCellsExamples.Utils.getSharedDataDir(DetectFileFormatOfEncryptedFiles.class) + "TechnicalArticles/";

		String filename = dataDir + "encryptedBook1.out.tmp";

       FileFormatInfo fileFormatInfo = FileFormatUtil.detectFileFormat(filename,"1234"); // The password is 1234

       if(fileFormatInfo.getFileFormatType() == FileFormatType.EXCEL_97_TO_2003) {
    	   System.out.println("File Format: EXCEL_97_TO_2003");
       } else if(fileFormatInfo.getFileFormatType() == FileFormatType.PPTX) {
    	   System.out.println("File Format: PPTX");
       } else if(fileFormatInfo.getFileFormatType() == FileFormatType.DOCX) {
    	   System.out.println("File Format: DOCX");
       }
       // ExEnd:1

	}
}
