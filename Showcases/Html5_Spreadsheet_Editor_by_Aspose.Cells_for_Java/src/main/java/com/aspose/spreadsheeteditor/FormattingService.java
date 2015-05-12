package com.aspose.spreadsheeteditor;

import java.io.Serializable;
import javax.faces.view.ViewScoped;
import javax.inject.Named;

/**
 *
 * @author Saqib Masood
 */
@Named(value = "formatting")
@ViewScoped
public class FormattingService implements Serializable {

    private boolean boldOptionEnabled;
    private boolean italicOptionEnabled;
    private boolean underlineOptionEnabled;
    private String fontSelectionOption;
    private int fontSizeOption;
    private String alignSelectionOption;
    private String fontColorSelectionOption;
    private String fillColorSelectionOption;

    public boolean isBoldOptionEnabled() {
        return boldOptionEnabled;
    }

    public void setBoldOptionEnabled(boolean boldOptionEnabled) {
        this.boldOptionEnabled = boldOptionEnabled;
    }

    public boolean isItalicOptionEnabled() {
        return italicOptionEnabled;
    }

    public void setItalicOptionEnabled(boolean italicOptionEnabled) {
        this.italicOptionEnabled = italicOptionEnabled;
    }

    public boolean isUnderlineOptionEnabled() {
        return underlineOptionEnabled;
    }

    public void setUnderlineOptionEnabled(boolean underlineOptionEnabled) {
        this.underlineOptionEnabled = underlineOptionEnabled;
    }

    public String getFontSelectionOption() {
        return fontSelectionOption;
    }

    public void setFontSelectionOption(String fontSelectionOption) {
        this.fontSelectionOption = fontSelectionOption;
    }

    public int getFontSizeOption() {
        return fontSizeOption;
    }

    public void setFontSizeOption(int fontSizeOption) {
        this.fontSizeOption = fontSizeOption;
    }

    public String getAlignSelectionOption() {
        return alignSelectionOption;
    }

    public void setAlignSelectionOption(String alignSelectionOption) {
        this.alignSelectionOption = alignSelectionOption;
    }

    public String getFontColorSelectionOption() {
        return this.fontColorSelectionOption;
    }

    public void setFontColorSelectionOption(String c) {
        this.fontColorSelectionOption = c;
    }

    public String getFillColorSelectionOption() {
        return this.fillColorSelectionOption;
    }

    public void setFillColorSelectionOption(String c) {
        this.fillColorSelectionOption = c;
    }

}

