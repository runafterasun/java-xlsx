package ru.objectsfill.reader;

import com.fasterxml.jackson.databind.ObjectMapper;
import org.dhatim.fastexcel.reader.CellAddress;
import ru.objectsfill.enums.BindingMode;
import ru.objectsfill.enums.LoopOperationType;
import ru.objectsfill.exception.ExcelImportException;

import ru.objectsfill.model.ExcelImportParamCore;
import ru.objectsfill.model.FieldParam;
import ru.objectsfill.reader.dto.TemplateDto;
import ru.objectsfill.reader.dto.TemplateEntryDto;
import ru.objectsfill.util.FieldNameUtils;

import java.io.InputStream;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

/**
 * Читает параметры шаблона из JSON-файла.
 *
 * <p>Формат JSON:
 * <pre>{@code
 * {
 *   "entries": [
 *     { "fieldName": "test.account",      "sheetName": "Sheet1", "cellAddress": {"row": 0, "col": 0} },
 *     { "fieldName": "for.dateList.rate", "sheetName": "Sheet1", "headerName": "Ставка %" }
 *   ]
 * }
 * }</pre>
 *
 * <p>Правила:
 * <ul>
 *   <li>{@code fieldName} и {@code sheetName} — обязательны для каждой записи.</li>
 *   <li>Хотя бы одно из {@code headerName} или {@code cellAddress} — обязательно.</li>
 *   <li>Если {@code headerName} присутствует — режим привязки {@link BindingMode#HEADER} (приоритет).</li>
 *   <li>Если только {@code cellAddress} — режим привязки {@link BindingMode#POSITION}.</li>
 * </ul>
 */
public class JsonTemplateReader implements TemplateReader {

    private static final ObjectMapper MAPPER = new ObjectMapper();

    private final InputStream jsonStream;

    /**
     * @param jsonStream поток JSON-файла с описанием маркеров шаблона
     */
    public JsonTemplateReader(InputStream jsonStream) {
        this.jsonStream = jsonStream;
    }

    /**
     * Парсит JSON, валидирует каждую запись, строит список {@link FieldParam}
     * и автоматически определяет {@link BindingMode} для каждого ключа в {@code importParam}.
     *
     * @param importParam параметры импорта с зарегистрированными объектами
     * @param out         список, в который добавляются построенные {@link FieldParam}
     * @throws ExcelImportException если JSON невалиден или нарушены обязательные правила
     */
    @Override
    public void read(ExcelImportParamCore importParam, List<FieldParam> out) {
        TemplateDto dto = parseJson();

        List<TemplateEntryDto> entries = dto.getEntries();
        if (entries == null || entries.isEmpty()) {
            throw new ExcelImportException("JSON template: 'entries' list is missing or empty");
        }

        for (int i = 0; i < entries.size(); i++) {
            validate(entries.get(i), i);
            out.add(toFieldParam(entries.get(i)));
        }

        applyBindingModes(entries, importParam);
    }

    /**
     * Десериализует JSON из потока в {@link TemplateDto}.
     *
     * @return разобранный DTO
     * @throws ExcelImportException если JSON невалиден
     */
    private TemplateDto parseJson() {
        try {
            return MAPPER.readValue(jsonStream, TemplateDto.class);
        } catch (Exception e) {
            throw new ExcelImportException("Failed to parse JSON template", e);
        }
    }

    /**
     * Проверяет обязательные поля записи.
     *
     * @param entry запись из JSON
     * @param index порядковый номер записи (для сообщения об ошибке)
     * @throws ExcelImportException если нарушено хотя бы одно правило
     */
    private static void validate(TemplateEntryDto entry, int index) {
        if (entry.getFieldName() == null || entry.getFieldName().isBlank()) {
            throw new ExcelImportException("JSON template entry[" + index + "]: 'fieldName' is required");
        }
        if (entry.getSheetName() == null || entry.getSheetName().isBlank()) {
            throw new ExcelImportException("JSON template entry[" + index + "]: 'sheetName' is required");
        }
        boolean hasHeader = entry.getHeaderName() != null && !entry.getHeaderName().isBlank();
        boolean hasCellAddress = entry.getCellAddress() != null;
        if (!hasHeader && !hasCellAddress) {
            throw new ExcelImportException(
                    "JSON template entry[" + index + "]: either 'headerName' or 'cellAddress' must be specified");
        }
    }

    /**
     * Конвертирует DTO в {@link FieldParam}.
     * Если {@code headerName} присутствует — устанавливает его (приоритет над {@code cellAddress}).
     * {@code cellAddress} устанавливается всегда, когда задан.
     *
     * @param entry валидированная запись из JSON
     * @return построенный {@link FieldParam}
     */
    private static FieldParam toFieldParam(TemplateEntryDto entry) {
        FieldParam fp = new FieldParam()
                .setFieldName(entry.getFieldName().trim())
                .setSheetName(entry.getSheetName());

        if (entry.getHeaderName() != null && !entry.getHeaderName().isBlank()) {
            fp.setHeaderName(entry.getHeaderName().trim());
        }
        if (entry.getCellAddress() != null) {
            fp.setCellAddress(new CellAddress(entry.getCellAddress().getRow(), entry.getCellAddress().getCol()));
        }
        return fp;
    }

    /**
     * Автоматически определяет {@link BindingMode} для каждого ключа в {@code importParam}
     * на основе записей JSON и устанавливает его в соответствующий {@link ru.objectsfill.model.ImportInformation}.
     *
     * <p>Правило:
     * <ul>
     *   <li>Если хотя бы одна запись группы содержит {@code headerName} — режим {@link BindingMode#HEADER}.</li>
     *   <li>Иначе — {@link BindingMode#POSITION}.</li>
     * </ul>
     *
     * @param entries     все записи JSON-шаблона
     * @param importParam параметры импорта
     */
    private static void applyBindingModes(List<TemplateEntryDto> entries, ExcelImportParamCore importParam) {
        Map<String, List<TemplateEntryDto>> byKey = entries.stream()
                .collect(Collectors.groupingBy(e -> FieldNameUtils.getKey(e.getFieldName())));

        byKey.forEach((key, group) -> {
            if (!importParam.getParamsMap().containsKey(key)) return;
            if (!key.startsWith(LoopOperationType.FOR.getType())) return;

            boolean hasHeader = group.stream()
                    .anyMatch(e -> e.getHeaderName() != null && !e.getHeaderName().isBlank());

            BindingMode mode = hasHeader ? BindingMode.HEADER : BindingMode.POSITION;
            importParam.getParamsMap().get(key).setBindingMode(mode);
        });
    }

}
