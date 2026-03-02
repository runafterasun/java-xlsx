# excel-read-write — FAQ

## Что это?

Библиотека для **импорта и экспорта** данных из/в Excel-файлы (`.xlsx`) по шаблону.
Шаблон описывает **где** находятся данные (через маркеры в ячейках).

Под капотом используется [fastexcel](https://github.com/dhatim/fastexcel) (чтение и запись данных)
и [Apache POI](https://poi.apache.org/) (потоковое чтение стилей из шаблона).

---

## Как подключить?

Опубликуйте библиотеку в Maven Local:

```bash
./gradlew publishToMavenLocal
```

Затем добавьте в ваш проект:

**Gradle:**
```groovy
repositories {
    mavenLocal()
    mavenCentral()
}

dependencies {
    implementation 'ru.objectsfill:excel-read-write:0.0.1-SNAPSHOT'
}
```

**Maven:**
```xml
<dependency>
    <groupId>ru.objectsfill</groupId>
    <artifactId>excel-read-write</artifactId>
    <version>0.0.1-SNAPSHOT</version>
</dependency>
```

---

## Как работают маркеры?

В **шаблонном** Excel-файле ячейки помечаются строками вида:

```
<ключ>.<поле>
```

| Маркер             | Значение                                                    |
|--------------------|-------------------------------------------------------------|
| `test.account`     | Одиночный объект с ключом `test`, поле `account`           |
| `for.dateList.rate`| Loop-список с ключом `for.dateList`, поле `rate`           |

- Ключи без префикса `for.` → **одиночный объект** (читается один раз).
- Ключи с префиксом `for.` → **список объектов** (читаются все строки под заголовком).

---

## Источники шаблонов

Маркеры могут быть описаны двумя способами.

### Excel-шаблон

Создайте `.xlsx`-файл, где в нужных ячейках стоят маркеры (`test.account`, `for.list.rate` и т. д.).
Строка **над** loop-маркером считается заголовком столбца.

```java
// tmpl — шаблонный файл с маркерами, data — файл с данными
ExcelImportUtil.importExcel(importParam, tmpl, data);
```

### JSON-шаблон

Альтернатива Excel-шаблону — JSON-файл, описывающий те же маркеры:

```json
{
  "entries": [
    { "fieldName": "test.account",       "sheetName": "Sheet1", "cellAddress": {"row": 3, "col": 0} },
    { "fieldName": "for.list.rate",      "sheetName": "Sheet1", "headerName": "RATE" },
    { "fieldName": "for.list.date",      "sheetName": "Sheet1", "headerName": "DATE" }
  ]
}
```

| Поле          | Обязателен | Описание                                          |
|---------------|-----------|---------------------------------------------------|
| `fieldName`   | да        | Маркер вида `ключ.поле`                           |
| `sheetName`   | да        | Имя листа                                         |
| `headerName`  | *         | Текст заголовка столбца (HEADER mode)             |
| `cellAddress` | *         | Координаты ячейки `{"row": N, "col": N}` (POSITION mode) |

\* Нужно хотя бы одно из двух. Если указан `headerName` — используется HEADER mode.

```java
var reader = new JsonTemplateReader(new FileInputStream("template.json"));
ExcelImportUtil.importExcel(importParam, reader, data);
```

---

## Режимы привязки колонок (BindingMode)

Применяется только к loop-блокам (`for.`-ключи).

| Режим      | Как находит колонку                                             |
|------------|----------------------------------------------------------------|
| `HEADER`   | Сканирует import-файл и ищет ячейку с текстом `headerName` (до 100 строк). **По умолчанию.** |
| `POSITION` | Берёт координаты маркера напрямую из шаблона, без поиска.      |

```java
// Явно задать режим (только при использовании Excel-шаблона;
// JSON-шаблон определяет режим автоматически)
importParam.getParamsMap().put("for.list",
    new ImportInformation()
        .setClazz(MyRow.class)
        .setBindingMode(BindingMode.POSITION));
```

---

## Как настроить импорт?

```java
// 1. Создаём контейнер параметров
var importParam = new ExcelImportParamCore();

// 2. Регистрируем объекты по ключам
importParam.getParamsMap().put("test",          new ImportInformation().setClazz(MyObject.class));
importParam.getParamsMap().put("for.dateList",  new ImportInformation().setClazz(MyRow.class));

// 3. Запускаем импорт
try (InputStream tmpl = ...; InputStream data = ...) {
    ExcelImportUtil.importExcel(importParam, tmpl, data);
}

// 4. Получаем результаты
MyObject obj = (MyObject) importParam.getParamsMap().get("test").getLoopLst().get(0);
List<Object> rows = importParam.getParamsMap().get("for.dateList").getLoopLst();
```

---

## Как выглядит шаблонный файл?

### Одиночный объект

| A               | B               |
|-----------------|-----------------|
| test.account    | test.offerDate  |

Данные читаются из **тех же координат** в import-файле.

### Loop-список

| A                    | B                    |
|----------------------|----------------------|
| Счёт                 | Дата                 |
| for.list.account     | for.list.date        |

- Строка **над** маркером — заголовок столбца (по нему находится колонка в import-файле).
- Все строки **под** заголовком в import-файле читаются в список.

---

## Какие типы данных поддерживаются?

| Тип ячейки Excel | Поведение                                     |
|-----------------|-----------------------------------------------|
| `NUMBER`        | Значение конвертируется в `String` через `BigDecimal.toPlainString()` |
| Любой другой    | Берётся `rawValue` ячейки как `String`        |
| Дата (`m` в формате) | Конвертируется в строку через `LocalDateTime.toString()` |

Целевые поля объектов должны быть `String` или `Integer` (для служебного поля `row`).

---

## Как добавить обязательные поля?

```java
var info = new ImportInformation()
        .setClazz(MyRow.class)
        .setRequiredFields(new ArrayList<>(List.of("account", "date")));

importParam.getParamsMap().put("for.list", info);
```

При обнаружении заголовка в import-файле поле удаляется из списка `requiredFields`.
После импорта можно проверить, что список пуст:

```java
List<String> missing = importParam.getParamsMap().get("for.list").getRequiredFields();
if (!missing.isEmpty()) {
    throw new IllegalStateException("Не найдены обязательные поля: " + missing);
}
```

---

## Предупреждения (warnings)

Некритичные проблемы (например, поле из шаблона отсутствует в целевом классе) не бросают исключение, а попадают в список предупреждений:

```java
ExcelImportUtil.importExcel(importParam, tmpl, data);

List<String> warnings = importParam.getWarnings();
if (!warnings.isEmpty()) {
    warnings.forEach(System.out::println);
}
```

---

## Как обрабатываются ошибки?

Критичные ошибки оборачиваются в специализированные исключения с контекстным сообщением:

| Исключение              | Когда бросается                  |
|-------------------------|----------------------------------|
| `ExcelImportException`  | Ошибка при чтении / импорте      |
| `ExcelExportException`  | Ошибка при записи / экспорте     |

Оба находятся в пакете `ru.objectsfill.exception` и наследуются от `RuntimeException`.

```java
try (InputStream tmpl = ...; InputStream data = ...) {
    ExcelImportUtil.importExcel(importParam, tmpl, data);
} catch (ExcelImportException e) {
    log.error("Ошибка импорта: {}", e.getMessage(), e);
}

try (InputStream tmpl = ...; OutputStream out = ...) {
    ExcelExportUtil.exportExcelWithStyles(exportParam, tmpl, out);
} catch (ExcelExportException e) {
    log.error("Ошибка экспорта: {}", e.getMessage(), e);
}
```

---

## Как настроить экспорт?

Экспорт использует тот же шаблон, что и импорт: маркеры в ячейках указывают, **куда** записывать данные.
Поддерживаются оба источника шаблонов — Excel-файл и JSON.

### Шаг 1 — Подготовьте классы данных

Классы должны иметь публичный конструктор без аргументов и getter-методы вида `getFieldName()`.
Те же классы, что используются при импорте, подходят и для экспорта без изменений.

```java
public class MyObject {
    private String account;
    private String offerDate;
    // getters / setters
}

public class MyRow {
    private String origin;
    private String destination;
    private String rate;
    // getters / setters
}
```

### Шаг 2 — Создайте контейнер параметров

```java
var exportParam = new ExcelExportParamCore();
```

### Шаг 3 — Зарегистрируйте данные по ключам шаблона

Ключи должны совпадать с теми, что используются в шаблоне (без имени поля после точки).

```java
// одиночный объект — ключ без префикса "for."
exportParam.getParamsMap().put("test",
    new ExportInformation().setDataList(List.of(myObject)));

// loop-список — ключ с префиксом "for."
exportParam.getParamsMap().put("for.dateList",
    new ExportInformation().setDataList(myRows));
```

> `setDataList` принимает `List<Object>`. Для одиночного объекта передайте список из одного элемента.

### Шаг 4 — Запустите экспорт

**Через Excel-шаблон:**

```java
try (InputStream tmpl = new FileInputStream("template.xlsx");
     OutputStream out = new FileOutputStream("result.xlsx")) {
    ExcelExportUtil.exportExcel(exportParam, tmpl, out);
}
```

**Через JSON-шаблон:**

```java
try (InputStream tmpl = new FileInputStream("template.json");
     OutputStream out = new FileOutputStream("result.xlsx")) {
    ExcelExportUtil.exportExcel(exportParam, new JsonTemplateReader(tmpl), out);
}
```

### Шаг 5 — Проверьте предупреждения

Некоторые ситуации не прерывают экспорт, а фиксируются в списке предупреждений:
- поле маркера отсутствует в классе данных;
- для маркера не задан `cellAddress` (поле пропускается).

```java
List<String> warnings = exportParam.getWarnings();
if (!warnings.isEmpty()) {
    warnings.forEach(System.out::println);
}
```

---

### Как выглядит результат (Вариант A)

Строка маркера заменяется **первой строкой данных**. Последующие строки записываются ниже.

```
Строка N−1:  | ORIGIN | DESTINATION | RATE |   ← заголовки (из headerName)
Строка N:    | RU     | DE          | 0.05 |   ← первый объект (на месте маркера)
Строка N+1:  | US     | FR          | 0.12 |   ← второй объект
Строка N+2:  | CN     | GB          | 0.08 |   ← третий объект
```

Если `headerName` не задан (POSITION mode), строка заголовков не записывается.

---

### Экспорт со стилями

Метод `exportExcelWithStyles` переносит стили ячеек одиночных объектов из шаблона в результат.
Apache POI читает шаблон **потоково** (SAX), не загружая весь файл в память — обрабатываются только первые N строк.

```java
// По умолчанию читаются стили из первых 100 строк шаблона
try (InputStream tmpl = new FileInputStream("template.xlsx");
     OutputStream out = new FileOutputStream("result.xlsx")) {
    ExcelExportUtil.exportExcelWithStyles(exportParam, tmpl, out);
}

// Кастомная глубина чтения стилей (если маркеры ниже 100-й строки)
try (InputStream tmpl = new FileInputStream("template.xlsx");
     OutputStream out = new FileOutputStream("result.xlsx")) {
    ExcelExportUtil.exportExcelWithStyles(exportParam, tmpl, out, 200);
}
```

| Что переносится               | Поддержка |
|-------------------------------|-----------|
| Жирный, курсив, подчёркивание | ✓         |
| Шрифт, размер, цвет шрифта    | ✓         |
| Цвет фона (заливка)           | ✓         |
| Границы (thin, medium, thick…)| ✓         |
| Числовой формат               | ✓         |
| Горизонтальное выравнивание   | ✓         |
| Перенос текста                | ✓         |
| Условное форматирование       | ✗         |
| Формулы                       | ✗         |

> Стили применяются только к ячейкам **одиночных объектов**.
> Табличные строки (loop) записываются без стилей — намеренно.

---

### Ограничения экспорта

- При использовании `exportExcel` (без стилей) создаётся файл с «голыми» значениями — без форматирования шаблона.
- `exportExcelWithStyles` работает только с Excel-шаблоном; JSON-шаблон стили не хранит.
- Если `cellAddress` не задан для маркера — поле пропускается с предупреждением.
- При экспорте с несколькими листами данные loop-блока записываются на **каждый** лист, где шаблон содержит маркеры этого блока.

---

## Ограничения

- Читается не более **100 строк** на лист для одиночных объектов.
- Поля целевых классов заполняются через `setter`-методы (формат `setFieldName`).
- Целевые классы должны иметь **публичный конструктор без аргументов**.
- Один loop-блок на лист (несколько `for.`-ключей на одном листе не поддерживается).
