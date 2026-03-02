package ru.objectsfill.style;

/**
 * Атрибуты стиля одной ячейки, извлечённые из Excel-шаблона.
 * Хранит только то, что поддерживает fastexcel-writer.
 * Не зависит от Apache POI или fastexcel — чистый DTO.
 */
public class CellStyleAttributes {

    private boolean bold;
    private boolean italic;
    private boolean underline;
    private boolean wrapText;

    private String fontName;
    private short  fontSize;
    private String fontColor;       // "RRGGBB" или null

    private String fillColor;       // "RRGGBB" или null
    private String numberFormat;    // например "#,##0.00" или null
    private String horizontalAlignment; // "left", "center", "right", ... или null

    private String borderTopStyle;
    private String borderBottomStyle;
    private String borderLeftStyle;
    private String borderRightStyle;

    public boolean isBold()      { return bold; }
    public boolean isItalic()    { return italic; }
    public boolean isUnderline() { return underline; }
    public boolean isWrapText()  { return wrapText; }

    public CellStyleAttributes setBold(boolean bold)           { this.bold = bold; return this; }
    public CellStyleAttributes setItalic(boolean italic)       { this.italic = italic; return this; }
    public CellStyleAttributes setUnderline(boolean underline) { this.underline = underline; return this; }
    public CellStyleAttributes setWrapText(boolean wrapText)   { this.wrapText = wrapText; return this; }

    public String getFontName()  { return fontName; }
    public short  getFontSize()  { return fontSize; }
    public String getFontColor() { return fontColor; }

    public CellStyleAttributes setFontName(String fontName)   { this.fontName = fontName; return this; }
    public CellStyleAttributes setFontSize(short fontSize)    { this.fontSize = fontSize; return this; }
    public CellStyleAttributes setFontColor(String fontColor) { this.fontColor = fontColor; return this; }

    public String getFillColor()            { return fillColor; }
    public String getNumberFormat()         { return numberFormat; }
    public String getHorizontalAlignment()  { return horizontalAlignment; }

    public CellStyleAttributes setFillColor(String fillColor)                       { this.fillColor = fillColor; return this; }
    public CellStyleAttributes setNumberFormat(String numberFormat)                 { this.numberFormat = numberFormat; return this; }
    public CellStyleAttributes setHorizontalAlignment(String horizontalAlignment)   { this.horizontalAlignment = horizontalAlignment; return this; }

    public String getBorderTopStyle()    { return borderTopStyle; }
    public String getBorderBottomStyle() { return borderBottomStyle; }
    public String getBorderLeftStyle()   { return borderLeftStyle; }
    public String getBorderRightStyle()  { return borderRightStyle; }

    public CellStyleAttributes setBorderTopStyle(String borderTopStyle)       { this.borderTopStyle = borderTopStyle; return this; }
    public CellStyleAttributes setBorderBottomStyle(String borderBottomStyle) { this.borderBottomStyle = borderBottomStyle; return this; }
    public CellStyleAttributes setBorderLeftStyle(String borderLeftStyle)     { this.borderLeftStyle = borderLeftStyle; return this; }
    public CellStyleAttributes setBorderRightStyle(String borderRightStyle)   { this.borderRightStyle = borderRightStyle; return this; }
}
