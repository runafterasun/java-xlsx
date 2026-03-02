package ru.objectsfill.style;

import org.dhatim.fastexcel.BorderSide;
import org.dhatim.fastexcel.BorderStyle;
import org.dhatim.fastexcel.StyleSetter;
import org.dhatim.fastexcel.Worksheet;

import java.math.BigDecimal;

/**
 * Применяет {@link CellStyleAttributes} к ячейке через fastexcel {@link StyleSetter}.
 */
public class FastexcelStyleApplier {

    private FastexcelStyleApplier() {}

    /**
     * Применяет стили из {@link CellStyleAttributes} к ячейке через fastexcel {@link StyleSetter}.
     * Если {@code attrs} равен {@code null} — метод ничего не делает.
     *
     * @param ws   лист, содержащий ячейку
     * @param row  индекс строки (0-based)
     * @param col  индекс столбца (0-based)
     * @param attrs атрибуты стиля, прочитанные из шаблона
     */
    public static void apply(Worksheet ws, int row, int col, CellStyleAttributes attrs) {
        if (attrs == null) return;

        StyleSetter style = ws.style(row, col);

        if (attrs.isBold())      style.bold();
        if (attrs.isItalic())    style.italic();
        if (attrs.isUnderline()) style.underlined();
        if (attrs.isWrapText())  style.wrapText(true);

        if (notBlank(attrs.getFontName()))  style.fontName(attrs.getFontName());
        if (attrs.getFontSize() > 0)        style.fontSize(BigDecimal.valueOf(attrs.getFontSize()));
        if (notBlank(attrs.getFontColor())) style.fontColor(attrs.getFontColor());

        if (notBlank(attrs.getFillColor()))  style.fillColor(attrs.getFillColor());

        if (notBlank(attrs.getNumberFormat())
                && !"General".equalsIgnoreCase(attrs.getNumberFormat())) {
            style.format(attrs.getNumberFormat());
        }

        if (notBlank(attrs.getHorizontalAlignment())) {
            style.horizontalAlignment(attrs.getHorizontalAlignment());
        }

        applyBorders(style, attrs);

        style.set();
    }

    /**
     * Применяет стили границ (top/bottom/left/right) к {@link StyleSetter}.
     * Неизвестные или {@code NONE}-значения пропускаются.
     *
     * @param style объект StyleSetter для текущей ячейки
     * @param attrs атрибуты стиля
     */
    private static void applyBorders(StyleSetter style, CellStyleAttributes attrs) {
        BorderStyle top    = mapBorder(attrs.getBorderTopStyle());
        BorderStyle bottom = mapBorder(attrs.getBorderBottomStyle());
        BorderStyle left   = mapBorder(attrs.getBorderLeftStyle());
        BorderStyle right  = mapBorder(attrs.getBorderRightStyle());

        if (top    != null) style.borderStyle(BorderSide.TOP,    top);
        if (bottom != null) style.borderStyle(BorderSide.BOTTOM, bottom);
        if (left   != null) style.borderStyle(BorderSide.LEFT,   left);
        if (right  != null) style.borderStyle(BorderSide.RIGHT,  right);
    }

    /**
     * Маппинг имён POI BorderStyle → fastexcel BorderStyle.
     * Имена enum совпадают (THIN, MEDIUM, THICK, DASHED, DOTTED, DOUBLE, HAIR…).
     * Неизвестные значения пропускаются (возвращает null).
     */
    /**
     * Конвертирует имя стиля границы POI в fastexcel {@link BorderStyle}.
     * Имена enum совпадают (THIN, MEDIUM, THICK, DASHED, DOTTED, DOUBLE, HAIR…).
     * Возвращает {@code null} для {@code NONE} или неизвестных значений.
     *
     * @param poiName имя enum POI BorderStyle
     * @return соответствующий fastexcel {@link BorderStyle} или {@code null}
     */
    private static BorderStyle mapBorder(String poiName) {
        if (poiName == null || "NONE".equals(poiName)) return null;
        try {
            return BorderStyle.valueOf(poiName);
        } catch (IllegalArgumentException ignored) {
            return null;
        }
    }

    private static boolean notBlank(String s) {
        return s != null && !s.isBlank();
    }
}
