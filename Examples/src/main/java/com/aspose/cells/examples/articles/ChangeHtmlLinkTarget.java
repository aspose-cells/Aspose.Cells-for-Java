package com.aspose.cells.examples.articles;

import com.aspose.cells.HtmlLinkTargetType;
import com.aspose.cells.HtmlSaveOptions;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class ChangeHtmlLinkTarget {

	public static void main(String[] args) throws Exception {

		String dataDir = Utils.getSharedDataDir(ChangeHtmlLinkTarget.class) + "articles/";
		String inputPath = dataDir + "Sample1.xlsx";
		String outputPath = dataDir + "CHLinkTarget.html";

		Workbook workbook = new Workbook(inputPath);

		HtmlSaveOptions opts = new HtmlSaveOptions();
		opts.setLinkTargetType(HtmlLinkTargetType.SELF);

		workbook.save(outputPath, opts);

		System.out.println("File saved " + outputPath);

	}
}
