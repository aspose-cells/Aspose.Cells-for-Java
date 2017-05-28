package AsposeCellsExamples.TechnicalArticles;

import com.aspose.cells.ExternalConnection;
import com.aspose.cells.WebQueryConnection;
import com.aspose.cells.Workbook;
import AsposeCellsExamples.Utils;

public class GetDataConnection {

	public static void main(String[] args) throws Exception {

		String dataDir = Utils.getSharedDataDir(GetDataConnection.class) + "TechnicalArticles/";
		String inputPath = dataDir + "WebQuerySample.xlsx";

		Workbook workbook = new Workbook(inputPath);

		ExternalConnection connection = workbook.getDataConnections().get(0);

		if (connection instanceof WebQueryConnection) {
			WebQueryConnection webQuery = (WebQueryConnection) connection;
			System.out.println("Web Query URL: " + webQuery.getUrl());
		}

	}
}
