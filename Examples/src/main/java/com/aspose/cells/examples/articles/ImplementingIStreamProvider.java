package com.aspose.cells.examples.articles;

import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class ImplementingIStreamProvider {
	public static void main(String[] args) throws Exception {

		String dataDir = Utils.getSharedDataDir(ImplementingIStreamProvider.class) + "articles/";
		Workbook wb = new Workbook(dataDir + "sample.xlsx");
		ImplementingIStreamProvider options = new ImplementingIStreamProvider();
		options.setStreamProvider(new ExportStreamProvider(dataDir));
		wb.save(dataDir + "IIStreamProvider-out.html");

	}

	private void setStreamProvider(ExportStreamProvider exportStreamProvider) {
		// TODO Auto-generated method stub
		
	}
}
