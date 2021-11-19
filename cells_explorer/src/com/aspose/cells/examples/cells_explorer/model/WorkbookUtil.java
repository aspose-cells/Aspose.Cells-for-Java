package com.aspose.cells.examples.cells_explorer.model;

import java.util.ArrayList;

import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.WorksheetCollection;

public class WorkbookUtil
{
	
	public static CellsNode loadExcelFile(String fileName) throws Exception
    {		
        Workbook wb = new Workbook(fileName);        
        
        CellsNode rootNode = new CellsNode();
        rootNode.setNodeName("Workbook");
        rootNode.setNodeContent(wb.getFileName());
        rootNode.setNodeType(CellsNodeType.ROOT_NODE);
        rootNode.setChildList(new ArrayList<CellsNode>()); 

        WorksheetCollection sheets = wb.getWorksheets();
        int sheetCount = sheets.getCount();

        CellsNode sheetsNode = new CellsNode();
        sheetsNode.setNodeName("Worksheets");
        sheetsNode.setNodeContent("Worksheet count: " + sheetCount);
        sheetsNode.setNodeType(CellsNodeType.STRUCTURE_NODE);
        sheetsNode.setChildList(new ArrayList<CellsNode>()); 

        rootNode.addChild(sheetsNode);
        
        for (int i = 0; i < sheetCount; i++)
        {
            addWorksheetNodes(sheets.get(i), sheetsNode);
        }
        
        return rootNode;
        
    }

    private static void addWorksheetNodes(Worksheet sheet, CellsNode parent)
    {
        CellsNode sheetNode = new CellsNode();
        sheetNode.setNodeName(sheet.getName());
        sheetNode.setNodeContent(CellsNodeContentUtil.getWorksheetNodeContent(sheet));
        sheetNode.setNodeType(CellsNodeType.STRUCTURE_NODE);
        sheetNode.setChildList(new ArrayList<CellsNode>());        
        parent.addChild(sheetNode);

        WorksheetUtil sheetUtil = new WorksheetUtil(sheet, sheetNode);
        sheetUtil.addWorksheetNode();            
    }
}
