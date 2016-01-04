/**
 * NOTICE: ORIGINAL FILE MODIFIED
 */

package com.aspose.cells.examples.featurescomparison.workbook.opensavespreadsheet;

import org.docx4j.openpackaging.io.SaveToZipFile;
import org.docx4j.openpackaging.packages.OpcPackage;

import com.aspose.cells.examples.Utils;

public class Xlsx4jOpenSaveSpreadSheet
{
    /**
     * @param args
     */
    public static void main(String[] args) throws Exception {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(Xlsx4jOpenSaveSpreadSheet.class);
        String inputfilepath  = dataDir + "pivot.xlsm";

        boolean save = true;
        String outputfilepath = dataDir + "pivot-rtt-xlsx4j.xlsm";

        // Open a document from the file system
        // 1. Load the Package
        OpcPackage pkg = OpcPackage.load(new java.io.File(inputfilepath));

        // Save it
        if (save) {		
            SaveToZipFile saver = new SaveToZipFile(pkg);
            saver.save(outputfilepath);
        }
    }
}
