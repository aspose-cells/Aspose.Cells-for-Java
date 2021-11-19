package com.aspose.cells.examples.cells_explorer.model;


import java.util.ArrayList;
import java.util.Iterator;

import com.aspose.cells.Cell;
import com.aspose.cells.CellArea;
import com.aspose.cells.CellsHelper;
import com.aspose.cells.Chart;
import com.aspose.cells.ChartCollection;
import com.aspose.cells.ChartShape;
import com.aspose.cells.Column;
import com.aspose.cells.ColumnCollection;
import com.aspose.cells.Comment;
import com.aspose.cells.CommentCollection;
import com.aspose.cells.CommentShape;
import com.aspose.cells.ConditionalFormattingCollection;
import com.aspose.cells.FormatCondition;
import com.aspose.cells.FormatConditionCollection;
import com.aspose.cells.GroupShape;
import com.aspose.cells.Hyperlink;
import com.aspose.cells.HyperlinkCollection;
import com.aspose.cells.ListObject;
import com.aspose.cells.ListObjectCollection;
import com.aspose.cells.OleObject;
import com.aspose.cells.OleObjectCollection;
import com.aspose.cells.Picture;
import com.aspose.cells.PictureCollection;
import com.aspose.cells.PivotTable;
import com.aspose.cells.PivotTableCollection;
import com.aspose.cells.Row;
import com.aspose.cells.RowCollection;
import com.aspose.cells.Shape;
import com.aspose.cells.ShapeCollection;
import com.aspose.cells.Validation;
import com.aspose.cells.ValidationCollection;
import com.aspose.cells.Worksheet;


class WorksheetUtil
{
    private Worksheet sheet;
    private CellsNode sheetNode;
    WorksheetUtil(Worksheet currSheet, CellsNode parentNode)
    {
        this.sheet = currSheet;
        this.sheetNode = parentNode;
    }

    void addWorksheetNode()
    {
        addColumnNodes(sheet, sheetNode);
        addRowNodes(sheet, sheetNode);

        addShapeNodes(sheet, sheetNode);
        addPictureNodes(sheet, sheetNode);
        addChartNodes(sheet, sheetNode);
        addCommentNodes(sheet, sheetNode);
        addOleObjectNodes(sheet, sheetNode);
        addHyperlinkNodes(sheet, sheetNode);
        addMergedAreasNode(sheet, sheetNode);
        addTablesNode(sheet, sheetNode);
        addPivotTablesNode(sheet, sheetNode);
        addValidationsNode(sheet, sheetNode);
        addConditionalFormattingsNode(sheet, sheetNode);
    }

    private void addConditionalFormattingsNode(Worksheet currSheet, CellsNode parent)
    {
        ConditionalFormattingCollection confitionalFormats = currSheet.getConditionalFormattings();
        int confitionalFormatsCount = confitionalFormats.getCount();
        if (confitionalFormatsCount > 0)
        {
            CellsNode conditionalFormatsNode = new CellsNode();
            conditionalFormatsNode.setNodeName("ConditionalFormattings");
            conditionalFormatsNode.setNodeContent("FormatConditionCollection count: " + confitionalFormatsCount);
            conditionalFormatsNode.setNodeType(CellsNodeType.STRUCTURE_NODE);
            conditionalFormatsNode.setChildList(new ArrayList<CellsNode>());

            parent.addChild(conditionalFormatsNode);

            for (int i = 0; i < confitionalFormatsCount; i++)
            {
                addFormatConditionsNode(confitionalFormats.get(i), conditionalFormatsNode, i + 1);

            }
        }
    }

    private void addFormatConditionsNode(FormatConditionCollection formats, CellsNode parent, int index)
    {
        int formatCount = formats.getCount();
        if (formatCount > 0)
        {
            CellsNode formatsNode = new CellsNode();
            formatsNode.setNodeName("FormatConditions " + index);
            formatsNode.setNodeContent(CellsNodeContentUtil.getFormatConditionsNodeContent(formats));
            formatsNode.setNodeType(CellsNodeType.STRUCTURE_NODE);
            formatsNode.setChildList(new ArrayList<CellsNode>());

            parent.addChild(formatsNode);

            for (int i = 0; i < formatCount; i++)
            {
                addFormatConditionNode(formats.get(i), formatsNode, i + 1);

            }
        }
    }

    private void addFormatConditionNode(FormatCondition format, CellsNode parent, int index)
    {
        CellsNode tableNode = new CellsNode();
        tableNode.setNodeName("FormatCondition " + index);
        tableNode.setNodeContent(CellsNodeContentUtil.getFormatConditionNodeContent(format));
        tableNode.setNodeType(CellsNodeType.LEAF_NODE);
        tableNode.setChildList(new ArrayList<CellsNode>());

        parent.addChild(tableNode);
    }

    private void addValidationsNode(Worksheet currSheet, CellsNode parent)
    {
        ValidationCollection validations = currSheet.getValidations();
        int validationCount = validations.getCount();
        if (validationCount > 0)
        {
            CellsNode validationsNode = new CellsNode();
            validationsNode.setNodeName("Validations");
            validationsNode.setNodeContent("Validation count: " + validationCount);
            validationsNode.setNodeType(CellsNodeType.STRUCTURE_NODE);
            validationsNode.setChildList(new ArrayList<CellsNode>());

            parent.addChild(validationsNode);

            for (int i = 0; i < validationCount; i++)
            {
                addValidationNode(validations.get(i), validationsNode, (i + 1));

            }
        }
    }

    private void addValidationNode(Validation validation, CellsNode parent, int index)
    {
        CellsNode validationNode = new CellsNode();
        validationNode.setNodeName("Validation " + index);
        validationNode.setNodeContent(CellsNodeContentUtil.getValidationNodeContent(validation));
        validationNode.setNodeType(CellsNodeType.LEAF_NODE);
        validationNode.setChildList(new ArrayList<CellsNode>());

        parent.addChild(validationNode);
    }


    private void addPivotTablesNode(Worksheet currSheet, CellsNode parent)
    {
        PivotTableCollection tables = currSheet.getPivotTables();
        int tableCount = tables.getCount();
        if (tableCount > 0)
        {
            CellsNode tablesNode = new CellsNode();
            tablesNode.setNodeName("PivotTables");
            tablesNode.setNodeContent("PivotTable count: " + tableCount);
            tablesNode.setNodeType(CellsNodeType.STRUCTURE_NODE);
            tablesNode.setChildList(new ArrayList<CellsNode>());

            parent.addChild(tablesNode);

            for (int i = 0; i < tableCount; i++)
            {
                addPivotTableNode(tables.get(i), tablesNode);

            }
        }
    }

    private void addPivotTableNode(PivotTable table, CellsNode parent)
    {
        CellsNode tableNode = new CellsNode();
        tableNode.setNodeName(table.getName());
        tableNode.setNodeContent(CellsNodeContentUtil.getPivotTableNodeContent(table));
        tableNode.setNodeType(CellsNodeType.LEAF_NODE);
        tableNode.setChildList(new ArrayList<CellsNode>());

        parent.addChild(tableNode);
    }

    private void addTablesNode(Worksheet currSheet, CellsNode parent)
    {
        ListObjectCollection tables = currSheet.getListObjects();
        int tableCount = tables.getCount();
        if (tableCount > 0)
        {
            CellsNode tablesNode = new CellsNode();
            tablesNode.setNodeName("Tables");
            tablesNode.setNodeContent("Table count: " + tableCount);
            tablesNode.setNodeType(CellsNodeType.STRUCTURE_NODE);
            tablesNode.setChildList(new ArrayList<CellsNode>());

            parent.addChild(tablesNode);

            for (int i = 0; i < tableCount; i++)
            {
                addTableNode(tables.get(i), tablesNode);

            }
        }
    }

    private void addTableNode(ListObject table, CellsNode parent)
    {
        CellsNode tableNode = new CellsNode();
        tableNode.setNodeName(table.getDisplayName());
        tableNode.setNodeContent(CellsNodeContentUtil.getTableNodeContent(table));
        tableNode.setNodeType(CellsNodeType.LEAF_NODE);
        tableNode.setChildList(new ArrayList<CellsNode>());

        parent.addChild(tableNode);
    }

    @SuppressWarnings("rawtypes")
	private void addMergedAreasNode(Worksheet currSheet, CellsNode parent)
    {
        ArrayList mergedCells = currSheet.getCells().getMergedCells();
        int mergedCellsCount = mergedCells.size();
        if (mergedCellsCount > 0)
        {
            CellsNode mergedCellsNode = new CellsNode();
            mergedCellsNode.setNodeName("MergedAreas");
            mergedCellsNode.setNodeContent("MergedArea count: " + mergedCellsCount);
            mergedCellsNode.setNodeType(CellsNodeType.STRUCTURE_NODE);
            mergedCellsNode.setChildList(new ArrayList<CellsNode>());

            parent.addChild(mergedCellsNode);

            for (int i = 0; i < mergedCellsCount; i++)
            {
                addMergedAreaNode((CellArea)mergedCells.get(i), mergedCellsNode);

            }
        }
    }

    private void addMergedAreaNode(CellArea area, CellsNode parent)
    {
        CellsNode oleNode = new CellsNode();

        String range = CellsNodeContentUtil.rangeToName(area.StartRow, area.StartColumn, area.EndRow, area.EndColumn);
        oleNode.setNodeName(range);
        oleNode.setNodeContent(range);
        oleNode.setNodeType(CellsNodeType.LEAF_NODE);
        oleNode.setChildList(new ArrayList<CellsNode>());

        parent.addChild(oleNode);
    }



    private void addHyperlinkNodes(Worksheet currSheet, CellsNode parent)
    {
        HyperlinkCollection hyperlinks = currSheet.getHyperlinks();
        int hyperlinkCount = hyperlinks.getCount();
        if (hyperlinkCount > 0)
        {
            CellsNode hyperlinksNode = new CellsNode();
            hyperlinksNode.setNodeName("Hyperlinks");
            hyperlinksNode.setNodeContent("Hyperlink count: " + hyperlinkCount);
            hyperlinksNode.setNodeType(CellsNodeType.STRUCTURE_NODE);
            hyperlinksNode.setChildList(new ArrayList<CellsNode>());

            parent.addChild(hyperlinksNode);

            for (int i = 0; i < hyperlinkCount; i++)
            {
                addHyperlinkNode(hyperlinks.get(i), hyperlinksNode);

            }
        }
    }

    private void addHyperlinkNode(Hyperlink hyperlink, CellsNode parent)
    {
        CellsNode oleNode = new CellsNode();
        oleNode.setNodeName(hyperlink.getTextToDisplay());
        oleNode.setNodeContent(CellsNodeContentUtil.getHyperlinkNodeContent(hyperlink));
        oleNode.setNodeType(CellsNodeType.LEAF_NODE);
        oleNode.setChildList(new ArrayList<CellsNode>());

        parent.addChild(oleNode);
    }


    private void addOleObjectNodes(Worksheet currSheet, CellsNode parent)
    {
        OleObjectCollection oles = currSheet.getOleObjects();
        int oleCount = oles.getCount();
        if (oleCount > 0)
        {
            CellsNode olesNode = new CellsNode();
            olesNode.setNodeName("OleObjects");
            olesNode.setNodeContent("OleObject count: " + oleCount);
            olesNode.setNodeType(CellsNodeType.STRUCTURE_NODE);
            olesNode.setChildList(new ArrayList<CellsNode>());

            parent.addChild(olesNode);

            for (int i = 0; i < oleCount; i++)
            {
                addOleObjectNode(oles.get(i), olesNode);

            }
        }
    }

    private void addOleObjectNode(OleObject oleObj, CellsNode parent)
    {
        CellsNode oleNode = new CellsNode();
        oleNode.setNodeName(oleObj.getName());
        oleNode.setNodeContent(CellsNodeContentUtil.getOLEObjectNodeContent(oleObj));
        oleNode.setNodeType(CellsNodeType.LEAF_NODE);
        oleNode.setChildList(new ArrayList<CellsNode>());

        parent.addChild(oleNode);
    }

    private void addCommentNodes(Worksheet currSheet, CellsNode parent)
    {
        CommentCollection comments = currSheet.getComments();
        int commentCount = comments.getCount();
        if (commentCount > 0)
        {
            CellsNode commentsNode = new CellsNode();
            commentsNode.setNodeName("Comments");
            commentsNode.setNodeContent("Comment count: " + commentCount);
            commentsNode.setNodeType(CellsNodeType.STRUCTURE_NODE);
            commentsNode.setChildList(new ArrayList<CellsNode>());

            parent.addChild(commentsNode);

            for (int i = 0; i < commentCount; i++)
            {
                addCommentNode(comments.get(i), commentsNode);

            }
        }
    }

    private void addCommentNode(Comment comment, CellsNode parent)
    {
        CellsNode commentNode = new CellsNode();
        commentNode.setNodeName(comment.getCommentShape().getName());
        commentNode.setNodeContent(CellsNodeContentUtil.getCommentNodeContent(comment));
        commentNode.setNodeType(CellsNodeType.LEAF_NODE);
        commentNode.setChildList(new ArrayList<CellsNode>());

        parent.addChild(commentNode);
    }

    @SuppressWarnings("rawtypes")
	private void addColumnNodes(Worksheet currSheet, CellsNode parent)
    {
        ColumnCollection cols = currSheet.getCells().getColumns();
        int colCount = cols.getCount();
        if (colCount > 0)
        {
            CellsNode colsNode = new CellsNode();
            colsNode.setNodeName("Columns");
            colsNode.setNodeContent("Column count: " + colCount);
            colsNode.setNodeType(CellsNodeType.STRUCTURE_NODE);
            colsNode.setChildList(new ArrayList<CellsNode>());
            parent.addChild(colsNode);

            Iterator iter = cols.iterator();
            while (iter.hasNext())
            {
                Column curr = (Column)iter.next();
                CellsNode currRowNode = getColumnNode(curr, colsNode);
                colsNode.addChild(currRowNode);
            }


        }
    }

    private CellsNode getColumnNode(Column col, CellsNode parent)
    {
        CellsNode colNode = new CellsNode();
        colNode.setNodeName("Column " + CellsHelper.columnIndexToName(col.getIndex()));
        colNode.setNodeContent(CellsNodeContentUtil.getColumnNodeContent(col));
        colNode.setNodeType(CellsNodeType.LEAF_NODE);
        colNode.setChildList(new ArrayList<CellsNode>());

        return colNode;
    }

    @SuppressWarnings("rawtypes")
	private void addRowNodes(Worksheet currSheet, CellsNode parent)
    {
        RowCollection rows = currSheet.getCells().getRows();
        int rowCount = rows.getCount();
        if (rowCount > 0)
        {
            CellsNode rowsNode = new CellsNode();
            rowsNode.setNodeName("Rows");
            rowsNode.setNodeContent("Row count: " + rowCount);
            rowsNode.setNodeType(CellsNodeType.STRUCTURE_NODE);
            rowsNode.setChildList(new ArrayList<CellsNode>());
            parent.addChild(rowsNode);

            Iterator iter = rows.iterator();
            while (iter.hasNext())
            {
                Row curr = (Row)iter.next();
                CellsNode currRowNode = getRowNode(curr, rowsNode);
                rowsNode.addChild(currRowNode);

                Iterator rowIter = curr.iterator();
                while (rowIter.hasNext())
                {
                    addCellNode((Cell)rowIter.next(), currRowNode);
                }
            }


        }
    }

    private CellsNode getRowNode(Row row, CellsNode parent)
    {
        CellsNode rowNode = new CellsNode();
        rowNode.setNodeName("Row " + (row.getIndex() + 1));
        rowNode.setNodeContent(CellsNodeContentUtil.getRowNodeContent(row));
        rowNode.setNodeType(CellsNodeType.STRUCTURE_NODE);
        rowNode.setChildList(new ArrayList<CellsNode>());

        return rowNode;
    }

    private void addCellNode(Cell currCell, CellsNode parent)
    {
        CellsNode cellNode = new CellsNode();
        cellNode.setNodeName(currCell.getName());
        cellNode.setNodeContent(CellsNodeContentUtil.getCellNodeContent(currCell));
        cellNode.setNodeType(CellsNodeType.LEAF_NODE);
        cellNode.setChildList(new ArrayList<CellsNode>());

        parent.addChild(cellNode);
    }

    private void addChartNodes(Worksheet currSheet, CellsNode parent)
    {
        ChartCollection charts = currSheet.getCharts();
        int chartCount = charts.getCount();
        if (chartCount > 0)
        {
            CellsNode chartsNode = new CellsNode();
            chartsNode.setNodeName("Charts");
            chartsNode.setNodeContent("Chart count: " + chartCount);
            chartsNode.setNodeType(CellsNodeType.STRUCTURE_NODE);
            chartsNode.setChildList(new ArrayList<CellsNode>());

            parent.addChild(chartsNode);

            for (int i = 0; i < chartCount; i++)
            {
                addChartNode(charts.get(i), chartsNode);

            }
        }
    }

    private void addChartNode(Chart chart, CellsNode parent)
    {
        CellsNode chartNode = new CellsNode();
        chartNode.setNodeName(chart.getName());
        chartNode.setNodeContent(CellsNodeContentUtil.getChartNodeContent(chart));
        chartNode.setNodeType(CellsNodeType.LEAF_NODE);
        chartNode.setChildList(new ArrayList<CellsNode>());

        parent.addChild(chartNode);
    }

    private void addPictureNodes(Worksheet currSheet, CellsNode parent)
    {
        PictureCollection picts = currSheet.getPictures();
        int pictsCount = picts.getCount();
        if (pictsCount > 0)
        {
            CellsNode pictsNode = new CellsNode();
            pictsNode.setNodeName("Pictures");
            pictsNode.setNodeContent("Picture count: " + pictsCount);
            pictsNode.setNodeType(CellsNodeType.STRUCTURE_NODE);
            pictsNode.setChildList(new ArrayList<CellsNode>());

            parent.addChild(pictsNode);

            for (int i = 0; i < pictsCount; i++)
            {
                addPictureNode(picts.get(i), pictsNode);

            }
        }
    }

    private void addPictureNode(Picture pict, CellsNode parent)
    {
        CellsNode pictNode = new CellsNode();
        pictNode.setNodeName(pict.getName());
        pictNode.setNodeContent(CellsNodeContentUtil.getPictureNodeContent(pict));
        pictNode.setNodeType(CellsNodeType.LEAF_NODE);
        pictNode.setChildList(new ArrayList<CellsNode>());

        parent.addChild(pictNode);
    }

    private void addShapeNodes(Worksheet currSheet, CellsNode parent)
    {
        ShapeCollection shapes = currSheet.getShapes();
        int shapeCount = shapes.getCount();
        if (shapeCount > 0)
        {
            CellsNode shapesNode = new CellsNode();
            shapesNode.setNodeName("Shapes");
            shapesNode.setNodeContent("Shape count: " + shapeCount);
            shapesNode.setNodeType(CellsNodeType.STRUCTURE_NODE);
            shapesNode.setChildList(new ArrayList<CellsNode>());

            parent.addChild(shapesNode);

            for (int i = 0; i < shapeCount; i++)
            {
                Shape shape = shapes.get(i);
                if (isNeedAddShape(shape))
                {
                    addShapeNode(shape, shapesNode, false);
                }

            }
        }

    }

    private void addShapeNode(Shape shape, CellsNode parent, boolean groupIn)
    {
        if (shape.getGroup() != null && !groupIn)
        {
            return;
        }
        boolean isGroupShape = shape instanceof GroupShape;

        CellsNode shapeNode = new CellsNode();
        shapeNode.setNodeName(shape.getName());
        shapeNode.setNodeContent(CellsNodeContentUtil.getShapeNodeContent(shape));

        if (isGroupShape)
        {
            shapeNode.setNodeType(CellsNodeType.STRUCTURE_NODE);
        }
        else
        {
            shapeNode.setNodeType(CellsNodeType.LEAF_NODE);
        }

        shapeNode.setChildList(new ArrayList<CellsNode>());

        parent.addChild(shapeNode);

        if (isGroupShape)
        {
            GroupShape groupShape = (GroupShape)shape;
            Shape[] children = groupShape.getGroupedShapes();
            int childCount = children.length;
            for (int i = 0; i < childCount; i++)
            {
                Shape child = children[i];
                addShapeNode(child, shapeNode, true);
            }
        }
    }

    private boolean isNeedAddShape(Shape shape)
    {
        boolean isPict = shape instanceof Picture;

        boolean isChart = shape instanceof ChartShape;

        boolean isOle = shape instanceof OleObject;

        boolean isComment = shape instanceof CommentShape;

        return !(isPict || isChart || isOle || isComment);
    }

}

