package com.aspose.cells.examples.articles;

import com.aspose.cells.HtmlLinkTargetType;
import com.aspose.cells.HtmlSaveOptions;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class ChangeHtmlLinkTarget {

    public static void main(String[] args)
            throws Exception {

        String dataDir = Utils.getDataDir(ChangeHtmlLinkTarget.class);
        String inputPath = dataDir + "Sample1.xlsx";
        String outputPath = dataDir + "Output.html";

        Workbook workbook = new Workbook(inputPath);

        HtmlSaveOptions opts = new HtmlSaveOptions();
        opts.setLinkTargetType(HtmlLinkTargetType.SELF);

        workbook.save(outputPath, opts);

        System.out.println("File saved " + outputPath);
    }
}
