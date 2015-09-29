package asposefeatures.datahandlingfeatures.tracingprecedentsanddependents.java;

import com.aspose.cells.Cell;
import com.aspose.cells.Cells;
import com.aspose.cells.CellsHelper;
import com.aspose.cells.ReferredArea;
import com.aspose.cells.ReferredAreaCollection;
import com.aspose.cells.Workbook;

public class TracingPrecedentsAndDependents
{
    public static void main(String[] args) throws Exception
    {
	String dataPath = "src/asposefeatures/datahandlingfeatures/tracingprecedentsanddependents/data/";

	// Instantiating a Workbook object
	Workbook workbook = new Workbook(dataPath + "workbook.xls");

	Cells cells = workbook.getWorksheets().get(0).getCells();
	Cell cell = cells.get("A12");

	// Tracing precedents of the cell A12.
	// The return array contains ranges and cells.
	ReferredAreaCollection ret = cell.getPrecedents();

	// Printing all the precedent cells' name.
	if (ret != null)
	{
	    for (int m = 0; m < ret.getCount(); m++)
	    {
		ReferredArea area = ret.get(m);
		StringBuilder stringBuilder = new StringBuilder();
		if (area.isExternalLink())
		{
		    stringBuilder.append("[");
		    stringBuilder.append(area.getExternalFileName());
		    stringBuilder.append("]");
		}
		stringBuilder.append(area.getSheetName());
		stringBuilder.append("!");
		stringBuilder.append(CellsHelper.cellIndexToName(area.getStartRow(),
			area.getStartColumn()));
		if (area.isArea())
		{
		    stringBuilder.append(":");
		    stringBuilder.append(CellsHelper.cellIndexToName(area.getEndRow(),
			    area.getEndColumn()));
		}
		System.out.println("Tracing Precedents: " + stringBuilder.toString());
	    }
	}

	// Get the A1 cell
	Cell c = cells.get("A5");
	// Get the all the Dependents of A5 cell
	Cell[] dependents = c.getDependents(true);
	for (int i = 0; i < dependents.length; i++)
	{
	    System.out.println("Tracing Dependents: " + dependents[i].getWorksheet().getName()
		    + dependents[i].getName() + ":" + dependents[i].getIntValue());
	}
    }
}
