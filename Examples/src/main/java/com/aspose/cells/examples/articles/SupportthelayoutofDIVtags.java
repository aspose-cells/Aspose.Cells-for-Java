package com.aspose.cells.examples.articles;

import java.io.ByteArrayInputStream;

import com.aspose.cells.HTMLLoadOptions;
import com.aspose.cells.LoadFormat;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.examples.Utils;

public class SupportthelayoutofDIVtags {
	public static void main(String[] args) throws Exception {

		// The path to the documents directory.
		String dataDir = Utils.getSharedDataDir(SupportthelayoutofDIVtags.class) + "articles/";

		// Html string
		String export_html = " <html> <body>     <table>         <tr>             <td>                 <div>This is some Text.</div>                 <div>                     <div>                         <span>This is some more Text</span>                     </div>                     <div>                         <span>abc@abc.com</span>                     </div>                     <div>                         <span>1234567890</span>                     </div>                     <div>                         <span>ABC DEF</span>                     </div>                 </div>                 <div>Generated On May 30, 2016 02:33 PM <br />Time Call Received from Jan 01, 2016 to May 30, 2016</div>             </td>             <td>                 <img src='ASpose_logo_100x100.png' />             </td>         </tr>     </table> </body> </html>";

		// Convert html string to byte array input stream
		byte[] bts = export_html.getBytes();
		ByteArrayInputStream bis = new ByteArrayInputStream(bts);

		// Specify HTML load options, support div tag layouts
		HTMLLoadOptions loadOptions = new HTMLLoadOptions(LoadFormat.HTML);
		loadOptions.setSupportDivTag(true);

		// Create workbook object from the html using load options
		Workbook wb = new Workbook(bis, loadOptions);

		// Auto fit rows and columns of first worksheet
		Worksheet ws = wb.getWorksheets().get(0);
		ws.autoFitRows();
		ws.autoFitColumns();

		// Save the workbook in xlsx format
		wb.save(dataDir + "SThelayoutofDIVtags_out.xlsx", SaveFormat.XLSX);

	}

}
