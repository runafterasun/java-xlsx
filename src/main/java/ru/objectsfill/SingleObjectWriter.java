package ru.objectsfill;

import org.dhatim.fastexcel.Worksheet;
import org.dhatim.fastexcel.reader.CellAddress;
import ru.objectsfill.model.FieldParam;
import ru.objectsfill.style.CellStyleAttributes;
import ru.objectsfill.style.FastexcelStyleApplier;
import ru.objectsfill.util.FieldNameUtils;
import ru.objectsfill.util.ObjectReader;

import java.util.List;
import java.util.Map;

class SingleObjectWriter {

    private SingleObjectWriter() {}

    /**
     * Записывает поля одиночного объекта в ячейки листа (без стилей).
     *
     * @param ws         лист для записи
     * @param params     параметры маркеров для данного листа
     * @param dataObject объект с данными
     * @param warnings   список предупреждений
     */
    static void write(Worksheet ws, List<FieldParam> params, Object dataObject, List<String> warnings) {
        write(ws, params, dataObject, null, warnings);
    }

    /**
     * Записывает поля одиночного объекта в ячейки листа с опциональным применением стилей.
     *
     * @param ws         лист для записи
     * @param params     параметры маркеров для данного листа
     * @param dataObject объект с данными
     * @param styles     стили ячеек шаблона (может быть {@code null} — тогда стили не применяются)
     * @param warnings   список предупреждений
     */
    static void write(Worksheet ws, List<FieldParam> params, Object dataObject,
                      Map<CellAddress, CellStyleAttributes> styles, List<String> warnings) {
        for (FieldParam fp : params) {
            if (fp.getCellAddress() == null) {
                warnings.add("No cellAddress for field '" + fp.getFieldName() + "', skipping");
                continue;
            }
            int row = fp.getCellAddress().getRow();
            int col = fp.getCellAddress().getColumn();
            String fieldName = FieldNameUtils.getFieldName(fp.getFieldName());
            Object value = ObjectReader.readField(dataObject, fieldName, warnings);
            writeValue(ws, row, col, value);
            if (styles != null) {
                FastexcelStyleApplier.apply(ws, row, col, styles.get(fp.getCellAddress()));
            }
        }
    }

    /**
     * Записывает значение в ячейку, выбирая тип (Number, Boolean, String) автоматически.
     * Если значение {@code null} — ячейка не изменяется.
     *
     * @param ws    лист для записи
     * @param row   индекс строки (0-based)
     * @param col   индекс столбца (0-based)
     * @param value записываемое значение
     */
    static void writeValue(Worksheet ws, int row, int col, Object value) {
        if (value == null) return;
        if (value instanceof Number num) {
            ws.value(row, col, num.doubleValue());
        } else if (value instanceof Boolean bool) {
            ws.value(row, col, bool);
        } else {
            ws.value(row, col, value.toString());
        }
    }
}
