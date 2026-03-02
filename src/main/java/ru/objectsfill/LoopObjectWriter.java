package ru.objectsfill;

import org.dhatim.fastexcel.Worksheet;
import ru.objectsfill.model.FieldParam;
import ru.objectsfill.util.FieldNameUtils;
import ru.objectsfill.util.ObjectReader;

import java.util.List;

class LoopObjectWriter {

    private LoopObjectWriter() {}

    /**
     * Записывает loop-данные на лист.
     *
     * Вариант A: строка маркера заменяется первой строкой данных.
     * Заголовки пишутся на строку выше маркера (если headerName задан).
     *
     * @param ws       лист для записи
     * @param params   параметры маркеров для данного листа
     * @param dataList список объектов для записи
     * @param warnings список предупреждений
     */
    static void write(Worksheet ws, List<FieldParam> params, List<Object> dataList, List<String> warnings) {
        writeHeaders(ws, params, warnings);
        writeRows(ws, params, dataList, warnings);
    }

    /**
     * Записывает строку заголовков на строку выше маркера (если {@code headerName} задан).
     * Если строка заголовка уходит в отрицательный индекс — добавляет предупреждение и пропускает.
     *
     * @param ws       лист для записи
     * @param params   параметры маркеров для данного листа
     * @param warnings список предупреждений
     */
    private static void writeHeaders(Worksheet ws, List<FieldParam> params, List<String> warnings) {
        for (FieldParam fp : params) {
            if (fp.getCellAddress() == null) continue;
            if (fp.getHeaderName() == null || fp.getHeaderName().isBlank()) continue;
            int headerRow = fp.getCellAddress().getRow() - 1;
            if (headerRow < 0) {
                warnings.add("Header row would be negative for field '" + fp.getFieldName() + "', skipping header");
                continue;
            }
            ws.value(headerRow, fp.getCellAddress().getColumn(), fp.getHeaderName());
        }
    }

    /**
     * Записывает строки данных начиная со строки маркера.
     * Каждый следующий объект пишется на одну строку ниже предыдущего.
     *
     * @param ws       лист для записи
     * @param params   параметры маркеров для данного листа
     * @param dataList список объектов для записи
     * @param warnings список предупреждений
     */
    private static void writeRows(Worksheet ws, List<FieldParam> params, List<Object> dataList, List<String> warnings) {
        for (int i = 0; i < dataList.size(); i++) {
            Object obj = dataList.get(i);
            for (FieldParam fp : params) {
                if (fp.getCellAddress() == null) continue;
                String fieldName = FieldNameUtils.getFieldName(fp.getFieldName());
                Object value = ObjectReader.readField(obj, fieldName, warnings);
                int dataRow = fp.getCellAddress().getRow() + i;
                SingleObjectWriter.writeValue(ws, dataRow, fp.getCellAddress().getColumn(), value);
            }
        }
    }
}
