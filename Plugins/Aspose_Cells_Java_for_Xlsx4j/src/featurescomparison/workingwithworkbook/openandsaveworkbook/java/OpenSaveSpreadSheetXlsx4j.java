/**
 * NOTICE: ORIGINAL FILE MODIFIED
 */

package featurescomparison.workingwithworkbook.openandsaveworkbook.java;

import org.docx4j.openpackaging.io.SaveToZipFile;
import org.docx4j.openpackaging.packages.OpcPackage;

public class OpenSaveSpreadSheetXlsx4j
{
	/**
	 * @param args
	 */
	public static void main(String[] args) throws Exception
	{
		String dataPath = "src/featurescomparison/workingwithworkbook/openandsaveworkbook/data/";
		
		String inputfilepath = dataPath + "pivot.xlsm";

		boolean save = true;
		String outputfilepath = dataPath + "pivot-rtt-xlsx4j.xlsm";

		// Open a document from the file system
		// 1. Load the Package
		OpcPackage pkg = OpcPackage.load(new java.io.File(inputfilepath));

		// Save it

		if (save)
		{
			SaveToZipFile saver = new SaveToZipFile(pkg);
			saver.save(outputfilepath);
			
			// Print Message
			System.out.println("Worksheet saved successfully.");
		}
	}
}
