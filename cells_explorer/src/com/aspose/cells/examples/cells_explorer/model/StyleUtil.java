package com.aspose.cells.examples.cells_explorer.model;

import com.aspose.cells.BackgroundType;
import com.aspose.cells.Border;
import com.aspose.cells.BorderType;
import com.aspose.cells.CellBorderType;
import com.aspose.cells.Color;
import com.aspose.cells.Font;
import com.aspose.cells.Style;
import com.aspose.cells.TextAlignmentType;


class StyleUtil
{
	public static boolean isNullOrEmpty(String str)
	{
		if(str==null ||str.length()==0 )
			return true;
		else
			return false;
		
	}
	
    static String getStyleContent(Style style)
    {
        StringBuilder builder = new StringBuilder();
        builder.append("\nNumber: " + style.getNumber());

        if (!isNullOrEmpty(style.getCultureCustom()))
        {
            builder.append("\nCultureCustom: " + style.getCultureCustom());
        }

        if (!isNullOrEmpty(style.getCustom()))
        {
            builder.append("\nCustom: " + style.getCustom());
        }

        if (!isNullOrEmpty(style.getInvariantCustom()))
        {
            builder.append("\nInvariantCustom: " + style.getInvariantCustom());
        }

        builder.append("\n" + getHTextAlignStr(style.getHorizontalAlignment()));
        builder.append("\n" + getVTextAlignStr(style.getVerticalAlignment()));
        builder.append("\nIndent: " + style.getIndentLevel());
        builder.append("\nRotationAngle: " + style.getRotationAngle());

        builder.append("\nPattern: " + getPatternStyleStr(style.getPattern()));
        builder.append("\nBackgroundColor: " + sysColorToRGBHexStr(style.getBackgroundColor()));
        builder.append("\nForegroundColor: " + sysColorToRGBHexStr(style.getForegroundColor()));

        builder.append("\nFont: " + getFontStr(style.getFont()));

        builder.append(getBordersStr(style));

        return builder.toString();
    }

    static String getPatternStyleStr(int val)
    {
        switch (val)
        {
            case BackgroundType.NONE:
                return "";
            case BackgroundType.SOLID:
                return "none";
            case BackgroundType.GRAY_75:
                return "gray-75";
            case BackgroundType.GRAY_50:
                return "gray-50";
            case BackgroundType.GRAY_25:
                return "gray-25";
            case BackgroundType.GRAY_12:
                return "gray-125";
            case BackgroundType.GRAY_6:
                return "gray-0625";
            case BackgroundType.HORIZONTAL_STRIPE:
                return "horz-stripe";
            case BackgroundType.VERTICAL_STRIPE:
                return "vert-stripe";
            case BackgroundType.REVERSE_DIAGONAL_STRIPE:
                return "reverse-diag-stripe";
            case BackgroundType.DIAGONAL_STRIPE:
                return "diag-stripe";
            case BackgroundType.DIAGONAL_CROSSHATCH:
                return "diag-cross";
            case BackgroundType.THICK_DIAGONAL_CROSSHATCH:
                return "thick-diag-cross";
            case BackgroundType.THIN_HORIZONTAL_STRIPE:
                return "thin-horz-stripe";
            case BackgroundType.THIN_VERTICAL_STRIPE:
                return "thin-vert-stripe";
            case BackgroundType.THIN_REVERSE_DIAGONAL_STRIPE:
                return "thin-reverse-diag-stripe";
            case BackgroundType.THIN_DIAGONAL_STRIPE:
                return "thin-diag-stripe";
            case BackgroundType.THIN_HORIZONTAL_CROSSHATCH:
                return "thin-horz-cross";
            case BackgroundType.THIN_DIAGONAL_CROSSHATCH:
                return "thin-diag-cross";
            default:
                return "";
        }
    }

    static String getFontStr(Font font)
    {
        StringBuilder builder = new StringBuilder();
        builder.append("\n{\nName:" + font.getName());
        builder.append("\nSize:" + font.getDoubleSize());
        String colorHex = sysColorToRGBHexStr(font.getColor());
        builder.append("\ncolor:" + colorHex);

        builder.append("\nIsBold:" + font.isBold());
        builder.append("\nIsItalic:" + font.isItalic());
        builder.append("\nIsStrikeout:" + font.isStrikeout());
        builder.append("\nIsSubscript:" + font.isSubscript());
        builder.append("\nIsSuperscript:" + font.isSuperscript());
        builder.append("\nUnderline:" + font.getUnderline());
        builder.append("\n}");
        return builder.toString();
    }

    static String getBordersStr(Style style)
    {
        String top = getBorderStr(style, BorderType.TOP_BORDER);
        StringBuilder builder = new StringBuilder();
        builder.append("\nborder-top:" + top);

        String right = getBorderStr(style, BorderType.RIGHT_BORDER);
        builder.append("\nborder-right:" + right);

        String bottom = getBorderStr(style, BorderType.BOTTOM_BORDER);
        builder.append("\nborder-bottom:" + bottom);

        String left = getBorderStr(style, BorderType.LEFT_BORDER);

        builder.append("\nborder-left:" + left);

        String diagonalDown = getBorderStr(style, BorderType.DIAGONAL_DOWN);
        builder.append("\nmso-diagonal-down:" + diagonalDown);

        String diagonalUp = getBorderStr(style, BorderType.DIAGONAL_UP);
        builder.append("\nmso-diagonal-up:" + diagonalUp);

        return builder.toString();
    }

    public static String getBorderStr(Style style, int borderType)
    {
        Border border = style.getBorders().getByBorderType(borderType);
        int lineStyle = border.getLineStyle();
        if (lineStyle == CellBorderType.NONE)
        {
            return "none";
        }
        StringBuilder borderStr = new StringBuilder(getBorderLineStyleStrPixel(lineStyle));

        borderStr.append(" #" + sysColorToRGBHexStr(border.getColor()));

        return borderStr.toString();
    }

    public static String getBorderLineStyleStrPixel(int val)
    {
        switch (val)
        {
            case CellBorderType.NONE:
                return "none";
            case CellBorderType.THIN:
                return "1px solid";
            case CellBorderType.MEDIUM:
                return "2px solid";
            case CellBorderType.DASHED:
                return "1px dashed";
            case CellBorderType.DOTTED:
                return "1px dotted";
            case CellBorderType.THICK:
                return "3px solid";
            case CellBorderType.DOUBLE:
                return "4px double";
            case CellBorderType.HAIR:
                return "1px hairline";
            case CellBorderType.MEDIUM_DASHED:
                return "2px dashed";
            case CellBorderType.DASH_DOT:
                return "1px dot-dash";
            case CellBorderType.MEDIUM_DASH_DOT:
                return "2px dot-dash";
            case CellBorderType.DASH_DOT_DOT:
                return "1px dot-dot-dash";
            case CellBorderType.MEDIUM_DASH_DOT_DOT:
                return "2px dot-dot-dash";
            case CellBorderType.SLANTED_DASH_DOT:
                return "2px dot-dash-slanted";
            default:
                return "";
        }
    }

    static String sysColorToRGBHexStr(Color color)
    {
        int val = color.toArgb();
        String s = String.format("%X", val);
        return s;
    }


    static String getHTextAlignStr(int val)
    {
        switch (val)
        {
            case TextAlignmentType.CENTER:
                return "text-align: center";
            case TextAlignmentType.DISTRIBUTED:
                return "text-align: distributed";
            case TextAlignmentType.JUSTIFY:
                return "text-align: justify";
            case TextAlignmentType.LEFT:
                return "text-align: left";
            case TextAlignmentType.RIGHT:
                return "text-align: right";
            case TextAlignmentType.CENTER_ACROSS:
                return "text-align: center-across";
            case TextAlignmentType.FILL:
                return "text-align: fill";
            case TextAlignmentType.GENERAL:
                return "text-align: general";
            default:
                return "";
        }
    }

    static String getVTextAlignStr(int val)
    {
        switch (val)
        {
            case TextAlignmentType.BOTTOM:
                return "vertical-align: bottom";
            case TextAlignmentType.CENTER:
                return "vertical-align: middle";
            case TextAlignmentType.DISTRIBUTED:
                return "vertical-align: distributed";
            case TextAlignmentType.JUSTIFY:
                return "vertical-align: justify";
            case TextAlignmentType.TOP:
                return "vertical-align: top";
            default:
                return "";
        }
    }
    
    
}

