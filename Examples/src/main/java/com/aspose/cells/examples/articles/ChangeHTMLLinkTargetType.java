package com.aspose.cells.examples.articles;

import com.aspose.cells.HtmlLinkTargetType;
import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class ChangeHTMLLinkTargetType {
	public static void main(String[] args) throws Exception {
		// ExStart:ChangeHTMLLinkTargetType
		// The path to the documents directory.
		String dataDir = Utils.getDataDir(ChangeHTMLLinkTargetType.class);
		Workbook workbook = new Workbook(dataDir + "source.xlsx");

		HtmlSaveOptions opts = new HtmlSaveOptions();
		opts.setLinkTargetType(HtmlLinkTargetType.SELF);

		workbook.save(dataDir + "out.html", opts);
		// ExEnd:ChangeHTMLLinkTargetType
	}
}
