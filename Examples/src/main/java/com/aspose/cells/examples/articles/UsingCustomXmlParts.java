package com.aspose.cells.examples.articles;

import com.aspose.cells.Workbook;
import com.aspose.cells.examples.Utils;

public class UsingCustomXmlParts {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(UsingCustomXmlParts.class) + "articles/";
		String booksXML = "<catalog><book><title>Complete C#</title><price>44</price></book><book><title>Complete Java</title><price>76</price></book><book><title>Complete SharePoint</title><price>55</price></book><book><title>Complete PHP</title><price>63</price></book><book><title>Complete VB.NET</title><price>72</price></book></catalog>";

		Workbook workbook = new Workbook();
		workbook.getContentTypeProperties().add("BookStore", booksXML);
		workbook.save(dataDir + "UsingCustomXmlParts_out.xlsx");

	}

}
