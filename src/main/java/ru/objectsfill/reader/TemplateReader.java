package ru.objectsfill.reader;

import ru.objectsfill.model.ExcelImportParamCore;
import ru.objectsfill.model.FieldParam;

import java.util.List;

/**
 * Стратегия чтения параметров шаблона.
 * Реализации определяют источник разметки: Excel-файл ({@link ExcelTemplateReader})
 * или JSON-файл ({@link JsonTemplateReader}).
 */
public interface TemplateReader {

    /**
     * Читает маркеры из источника шаблона и добавляет их в список {@code out}.
     * При необходимости обновляет {@link ru.objectsfill.enums.BindingMode}
     * в записях {@code importParam}.
     *
     * @param importParam параметры импорта с зарегистрированными объектами
     * @param out         список, в который добавляются найденные {@link FieldParam}
     */
    void read(ExcelImportParamCore importParam, List<FieldParam> out);
}
