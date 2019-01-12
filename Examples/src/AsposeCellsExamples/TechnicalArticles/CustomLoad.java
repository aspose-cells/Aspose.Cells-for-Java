package AsposeCellsExamples.TechnicalArticles;

import com.aspose.cells.LoadDataFilterOptions;
import com.aspose.cells.LoadFilter;
import com.aspose.cells.Worksheet;

public class CustomLoad extends LoadFilter
{
	public void startSheet(Worksheet sheet)
    {
        if (sheet.getName() == "Sheet2")
        {
            // Load everything from worksheet "Sheet2"
            this.setLoadDataFilterOptions(LoadDataFilterOptions.ALL);
        }
        else
        {
            // Load nothing
            this.setLoadDataFilterOptions(~LoadDataFilterOptions.ALL);
        }
    }
}
