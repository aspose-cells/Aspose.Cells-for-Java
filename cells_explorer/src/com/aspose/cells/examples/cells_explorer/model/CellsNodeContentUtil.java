package com.aspose.cells.examples.cells_explorer.model;

import com.aspose.cells.Cell;
import com.aspose.cells.CellArea;
import com.aspose.cells.CellValueType;
import com.aspose.cells.Chart;
import com.aspose.cells.ColorScale;
import com.aspose.cells.Column;
import com.aspose.cells.Comment;
import com.aspose.cells.DataBar;
import com.aspose.cells.DataBarFillType;
import com.aspose.cells.FormatCondition;
import com.aspose.cells.FormatConditionCollection;
import com.aspose.cells.FormatConditionType;
import com.aspose.cells.Hyperlink;
import com.aspose.cells.IconSet;
import com.aspose.cells.IconSetType;
import com.aspose.cells.ListObject;
import com.aspose.cells.OleObject;
import com.aspose.cells.OperatorType;
import com.aspose.cells.Picture;
import com.aspose.cells.PivotTable;
import com.aspose.cells.Row;
import com.aspose.cells.Shape;
import com.aspose.cells.TextDirectionType;
import com.aspose.cells.Validation;
import com.aspose.cells.ValidationType;
import com.aspose.cells.Worksheet;


class CellsNodeContentUtil
{
    static String getWorksheetNodeContent(Worksheet sheet)
    {
        StringBuilder builder = new StringBuilder();
        builder.append("Name: " + sheet.getName());
        builder.append("\nHidden: " + !sheet.isVisible());

        return builder.toString();
    }

    static String getRowNodeContent(Row row)
    {
        StringBuilder builder = new StringBuilder();
        builder.append("Index: " + row.getIndex());
        builder.append("\nHidden: " + row.isHidden());

        builder.append("\nHeight: " + row.getHeight() + "pt");

        builder.append(StyleUtil.getStyleContent(row.getStyle()));
        return builder.toString();
    }

    static String getColumnNodeContent(Column col)
    {
        StringBuilder builder = new StringBuilder();
        builder.append("Index: " + col.getIndex());
        builder.append("\nHidden: " + col.isHidden());

        builder.append("\nWidth: " + col.getWidth() + "pt");

        builder.append(StyleUtil.getStyleContent(col.getStyle()));
        return builder.toString();
    }

    static String getCellNodeContent(Cell cell)
    {
        StringBuilder builder = new StringBuilder();
        builder.append("Name: " + cell.getName());
        builder.append("\nValue: " + cell.getStringValue());
        builder.append("\nValueType: " + cellValueTypeToStr(cell.getType()));
        builder.append("\nIsFormula: " + cell.isFormula());
        builder.append("\nIsArrayFormula: " + cell.isArrayFormula());
        builder.append("\nIsSharedFormula: " + cell.isSharedFormula());
        builder.append("\nIsTableFormula: " + cell.isTableFormula());
        
        builder.append(StyleUtil.getStyleContent(cell.getStyle()));
        return builder.toString();
    }

    static String cellValueTypeToStr(int type)
    {
        switch (type)
        {
            case CellValueType.IS_BOOL:
                return "bool";
            case CellValueType.IS_DATE_TIME:
                return "datetime";
            case CellValueType.IS_ERROR:
                return "error";
            case CellValueType.IS_NULL:
                return "null";
            case CellValueType.IS_NUMERIC:
                return "numeric";
            case CellValueType.IS_STRING:
                return "string";
            default:
                return "unknown";
        }
    }

    static String cellAreaToName(CellArea area)
    {
        return rangeToName(area.StartRow, area.StartColumn, area.EndRow, area.EndColumn);
    }
    

    static String rangeToName(int startRow, int startColumn, int endRow, int endColumn)
    {
        StringBuilder sb = new StringBuilder();
        columnIndexToName(sb, startColumn);
        sb.append(startRow + 1);
        if (startRow != endRow
                 || startColumn != endColumn)
        {
            sb.append(":");
            columnIndexToName(sb, endColumn);
            sb.append(endRow + 1);

        }
        return sb.toString();
    }

    static void columnIndexToName(StringBuilder sb, int column)
    {
        if (column < 0)
        {
            throw new IllegalArgumentException();
        }
        else if (column < 26)
        {
            sb.append((char)(column + 'A'));
        }
        else if (column < 702)
        {
            char c1 = (char)((column % 26) + 'A');

            column = column / 26 - 1;
            sb.append((char)(column + 'A'));
            sb.append(c1);
        }
        else if (column <= 16383)
        {
            char c1 = (char)((column % 26) + 'A');
            column = column / 26 - 1;
            char c2 = (char)((column % 26) + 'A');
            column = column / 26 - 1;
            sb.append((char)(column + 'A'));
            sb.append(c2);
            sb.append(c1);

        }
        else
        {
            throw new IllegalArgumentException();
        }
    }
    static String getFormatConditionsNodeContent(FormatConditionCollection formats)
    {
        StringBuilder builder = new StringBuilder();
        int rangeCount = formats.getRangeCount();

        builder.append("FormatCondition count: " + formats.getCount());

        builder.append("\nRange:");
        for (int i = 0; i < rangeCount; i++)
        {
            CellArea area = formats.getCellArea(i);
            builder.append("\n" + cellAreaToName(area));
        }
        return builder.toString();
    }

    static String getFormatConditionNodeContent(FormatCondition format)
    {
        StringBuilder builder = new StringBuilder();

        String type = getFormatConditionName(format.getType());
        if (!StyleUtil.isNullOrEmpty(type))
        {
            builder.append("FormatConditionType: " + type);
        }

        String operType = conditionOperatorToString(format.getOperator());
        if (!StyleUtil.isNullOrEmpty(operType))
        {
            builder.append("\nOperator Type: " + operType);
        }
        if (!StyleUtil.isNullOrEmpty(format.getFormula1()))
        {
            builder.append("\nFormula1: " + format.getFormula1());
        }
        if (!StyleUtil.isNullOrEmpty(format.getFormula2()))
        {
            builder.append("\nFormula1: " + format.getFormula2());
        }

        IconSet iconset = format.getIconSet();
        if (iconset != null && format.getType() == FormatConditionType.ICON_SET)
        {
            builder.append("\nIconSet:");
            String iconsetType = iconSetTypeToString(iconset.getType());
            if (!StyleUtil.isNullOrEmpty(iconsetType))
            {
                builder.append("\nIconSetType: " + iconsetType);
            }
        }

        ColorScale colorscale = format.getColorScale();
        if (colorscale != null && format.getType() == FormatConditionType.COLOR_SCALE)
        {
            builder.append("\nColorScale:");
            if (colorscale.getMaxColor() != null)
            {
                builder.append("\nMaxColor: 0x" + StyleUtil.sysColorToRGBHexStr(colorscale.getMaxColor()));
            }
            if (colorscale.getMidColor() != null)
            {
                builder.append("\nMidColor: 0x" + StyleUtil.sysColorToRGBHexStr(colorscale.getMidColor()));
            }
            if (colorscale.getMinColor() != null)
            {
                builder.append("\nMinColor: 0x" + StyleUtil.sysColorToRGBHexStr(colorscale.getMinColor()));
            }
        }

        DataBar databar = format.getDataBar();

        if (databar != null && format.getType() == FormatConditionType.DATA_BAR)
        {
            builder.append("\nDataBar:");
            if (databar.getAxisColor() != null)
            {
                builder.append("\nAxisColor: 0x" + StyleUtil.sysColorToRGBHexStr(databar.getAxisColor()));
            }
            if (databar.getColor() != null)
            {
                builder.append("\nColor: 0x" + StyleUtil.sysColorToRGBHexStr(databar.getColor()));
            }

            String direction = dataBarTextDirectionTypeToString(databar.getDirection());
            if (!StyleUtil.isNullOrEmpty(direction))
            {
                builder.append("\nDirection: " + direction);
            }

            if (databar.getBarFillType() == DataBarFillType.GRADIENT)
            {
                builder.append("\nBarFillType: " + "Gradient");
            }
            else
            {
                builder.append("\nBarFillType: " + "Solid");
            }
        }

        if (format.getStyle() != null)
        {                
            StyleUtil.getStyleContent(format.getStyle());
        }
        
        
      
        return builder.toString();
    }


    static String dataBarTextDirectionTypeToString(int type)
    {
        switch (type)
        {
            case TextDirectionType.LEFT_TO_RIGHT:
                return "leftToRight";
            case TextDirectionType.RIGHT_TO_LEFT:
                return "rightToLeft";
            default:
                return null;
        }
    }

    static String iconSetTypeToString(int type)
    {
        switch (type)
        {
            case IconSetType.ARROWS_3:
                return "3Arrows";
            case IconSetType.ARROWS_4:
                return "4Arrows";
            case IconSetType.ARROWS_5:
                return "5Arrows";
            case IconSetType.ARROWS_GRAY_3:
                return "3ArrowsGray";
            case IconSetType.ARROWS_GRAY_4:
                return "4ArrowsGray";
            case IconSetType.ARROWS_GRAY_5:
                return "5ArrowsGray";
            case IconSetType.FLAGS_3:
                return "3Flags";
            case IconSetType.QUARTERS_5:
                return "5Quarters";
            case IconSetType.RATING_4:
                return "4Rating";
            case IconSetType.RATING_5:
                return "5Rating";
            case IconSetType.RED_TO_BLACK_4:
                return "4RedToBlack";
            case IconSetType.SIGNS_3:
                return "3Signs";
            case IconSetType.SYMBOLS_3:
                return "3Symbols";
            case IconSetType.SYMBOLS_32:
                return "3Symbols2";
            case IconSetType.TRAFFIC_LIGHTS_31:
                return "3TrafficLights1";
            case IconSetType.TRAFFIC_LIGHTS_32:
                return "3TrafficLights2";
            case IconSetType.TRAFFIC_LIGHTS_4:
                return "4TrafficLights";
            case IconSetType.STARS_3:
                return "3Stars";
            case IconSetType.BOXES_5:
                return "5Boxes";
            case IconSetType.TRIANGLES_3:
                return "3Triangles";
            case IconSetType.NONE:
                return "NoIcons";
            case IconSetType.SMILIES_3:
                return "3Smilies";
            case IconSetType.COLOR_SMILIES_3:
                return "3ColorSmilies";
            default:
                return "";
        }
    }

    static String conditionOperatorToString(int type)
    {
        switch (type)
        {
            case OperatorType.BETWEEN:
                return "Between";
            case OperatorType.EQUAL:
                return "Equal";
            case OperatorType.GREATER_OR_EQUAL:
                return "GreaterOrEqual";
            case OperatorType.GREATER_THAN:
                return "greater";
            case OperatorType.LESS_OR_EQUAL:
                return "LessOrEqual";
            case OperatorType.LESS_THAN:
                return "less";
            case OperatorType.NOT_BETWEEN:
                return "notBetween";
            case OperatorType.NOT_EQUAL:
                return "notEqual";
            default:
                return "";
        }
    }

    static String getFormatConditionName(int type)
    {
        switch (type)
        {
            case FormatConditionType.CELL_VALUE:
                return "cellIs";
            case FormatConditionType.EXPRESSION:
                return "expression";
            case FormatConditionType.COLOR_SCALE:
                return "colorScale";
            case FormatConditionType.DATA_BAR:
                return "dataBar";
            case FormatConditionType.ICON_SET:
                return "iconSet";
            case FormatConditionType.ABOVE_AVERAGE:
                return "aboveAverage";
            case FormatConditionType.BEGINS_WITH:
                return "beginsWith";
            case FormatConditionType.CONTAINS_BLANKS:
                return "containsBlanks";
            case FormatConditionType.CONTAINS_ERRORS:
                return "containsErrors";
            case FormatConditionType.CONTAINS_TEXT:
                return "containsText";
            case FormatConditionType.DUPLICATE_VALUES:
                return "duplicateValues";
            case FormatConditionType.ENDS_WITH:
                return "endsWith";
            case FormatConditionType.NOT_CONTAINS_BLANKS:
                return "notContainsBlanks";
            case FormatConditionType.NOT_CONTAINS_ERRORS:
                return "notContainsErrors";
            case FormatConditionType.NOT_CONTAINS_TEXT:
                return "notContainsText";
            case FormatConditionType.TIME_PERIOD:
                return "timePeriod";
            case FormatConditionType.TOP_10:
                return "top10";
            case FormatConditionType.UNIQUE_VALUES:
                return "uniqueValues";
            default:
                return "";
        }

    }

    static String getValidationNodeContent(Validation validation)
    {
        StringBuilder builder = new StringBuilder();

        String validationTypeStr = validationTypeToString(validation.getType());
        if (!StyleUtil.isNullOrEmpty(validationTypeStr))
        {
            builder.append("ValidationType: " + validationTypeStr);
        }

        CellArea[]  areas = validation.getAreas();
        builder.append("\nAreas: ");
        int len = areas.length;
        if (areas != null && areas.length > 0)
        {
            for (int i=0; i <len;i++)
            {
                builder.append("\n" + cellAreaToName(areas[i]));
            }
            
        }

        String operType = conditionOperatorToString(validation.getOperator());
        if (!StyleUtil.isNullOrEmpty(operType))
        {
            builder.append("\nOperator Type: " + operType);
        }

        if (validation.getValue1() != null)
        {
            builder.append("\nValue1: " + validation.getValue1());
        }

        if (validation.getValue2() != null)
        {
            builder.append("\nValue2: " + validation.getValue2());
        }

        

        return builder.toString();
    }

    static String validationTypeToString(int type)
    {
        switch (type)
        {
            case ValidationType.CUSTOM:
                return "custom";
            case ValidationType.ANY_VALUE:
                return "none";
            case ValidationType.DATE:
                return "date";
            case ValidationType.DECIMAL:
                return "decimal";
            case ValidationType.LIST:
                return "list";
            case ValidationType.TEXT_LENGTH:
                return "textLength";
            case ValidationType.TIME:
                return "time";
            case ValidationType.WHOLE_NUMBER:
                return "whole";
            default:
                return "";
        }
    }

    static String getPivotTableNodeContent(PivotTable table)
    {
        StringBuilder builder = new StringBuilder();
        builder.append("Name: " + table.getName());
        builder.append("\nPivotTable Range(include page fields): " + cellAreaToName(table.getTableRange2()));
        builder.append("\nDataRange: " + cellAreaToName(table.getTableRange1()));
        if (!StyleUtil.isNullOrEmpty(table.getPivotTableStyleName()))
        {
            builder.append("\nPivotTableStyleName: " + table.getPivotTableStyleName());
        }

        return builder.toString();
    }

    static String getTableNodeContent(ListObject table)
    {
        StringBuilder builder = new StringBuilder();
        builder.append("Name: " + table.getDisplayName());
   
        builder.append("\nDataRange: " + table.getDataRange().getAddress());
        if (!StyleUtil.isNullOrEmpty(table.getTableStyleName()))
        {
             builder.append("\nTableStyleName: " + table.getTableStyleName());                  
        }
       
        return builder.toString();
    }

    static String getHyperlinkNodeContent(Hyperlink hyperlink)
    {
        StringBuilder builder = new StringBuilder();
        builder.append("Address: " + hyperlink.getAddress());
        builder.append("\nArea: ");
        builder.append(hyperlink.getArea().toString());

        return builder.toString();
    }

    static String getOLEObjectNodeContent(OleObject oleObj)
    {
        StringBuilder builder = new StringBuilder();
        getCommenProperties(builder, oleObj);

        return builder.toString();
    }

    static String getCommentNodeContent(Comment comment)
    {
        StringBuilder builder = new StringBuilder();
        getCommenProperties(builder, comment.getCommentShape());

        return builder.toString(); 
    }

    static String getChartNodeContent(Chart chart)
    {
        StringBuilder builder = new StringBuilder();
        getCommenProperties(builder, chart.getChartObject());

        return builder.toString();
    }

    static String getPictureNodeContent(Picture shape)
    {
        StringBuilder builder = new StringBuilder();
        getCommenProperties(builder, shape);

        return builder.toString();     
    }


    static String getShapeNodeContent(Shape shape)
    {
        StringBuilder builder = new StringBuilder();
        getCommenProperties(builder, shape);            

        return builder.toString();
    }

    private static void getCommenProperties(StringBuilder builder, Shape shape)
    {
        builder.append("Name: " + shape.getName());
        builder.append("\nUpperLeftRow: ");
        builder.append(shape.getUpperLeftRow());
        builder.append("\nUpperLeftColumn: ");
        builder.append(shape.getUpperLeftColumn());

        builder.append("\nX: ");
        builder.append(shape.getX() + "px   ");
        builder.append("Y: ");
        builder.append(shape.getY() + "px");

        builder.append("\nWidth: ");
        builder.append(shape.getWidth() + "px   ");
        builder.append("Height: ");
        builder.append(shape.getHeight() + "px");
    }

}

