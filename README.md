# excel-read-write — FAQ

## How to use
From maven central <br>
[![Latest Release](https://maven-badges.sml.io/sonatype-central/ru.objectsfill/excel-read-write/badge.svg?subject=Latest%20Release&color=blue)](https://central.sonatype.com/artifact/ru.objectsfill/excel-read-write)
![Build Status](https://github.com/runafterasun/objects-fill-processor/actions/workflows/gradle.yml/badge.svg)

```xml
<dependency>
    <groupId>ru.objectsfill</groupId>
    <artifactId>excel-read-write</artifactId>
    <version>x.x.x</version>
</dependency>
```

```gradle
testImplementation 'ru.objectsfill:excel-read-write:x.x.x'
testAnnotationProcessor 'ru.objectsfill:excel-read-write:x.x.x'
```

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

| Маркер              | Значение                                                      |
|---------------------|---------------------------------------------------------------|
| `test.company`      | Одиночный объект с ключом `test`, поле `company`             |
| `for.user.salary`   | Loop-список с ключом `for.user`, поле `salary`               |

- Ключи без префикса `for.` → **одиночный объект** (читается один раз).
- Ключи с префиксом `for.` → **список объектов** (читаются все строки под заголовком).

---

## Источники шаблонов

Маркеры могут быть описаны двумя способами.

### Excel-шаблон

Создайте `.xlsx`-файл, где в нужных ячейках стоят маркеры (`test.company`, `for.user.salary` и т. д.).
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
    { "fieldName": "test.company",     "sheetName": "List1", "cellAddress": {"row": 3, "col": 0} },
    { "fieldName": "for.user.salary",  "sheetName": "List1", "headerName": "SALARY" },
    { "fieldName": "for.user.age",     "sheetName": "List1", "headerName": "AGE" }
  ]
}
```

| Поле          | Обязателен | Описание                                                  |
|---------------|-----------|-----------------------------------------------------------|
| `fieldName`   | да        | Маркер вида `ключ.поле`                                   |
| `sheetName`   | да        | Имя листа                                                 |
| `headerName`  | *         | Текст заголовка столбца (HEADER mode)                     |
| `cellAddress` | *         | Координаты ячейки `{"row": N, "col": N}` (POSITION mode) |

\* Нужно хотя бы одно из двух. Если указан `headerName` — используется HEADER mode.

```java
var reader = new JsonTemplateReader(new FileInputStream("template.json"));
ExcelImportUtil.importExcel(importParam, reader, data);
```

---

## Режимы привязки колонок (BindingMode)

Применяется только к loop-блокам (`for.`-ключи).

| Режим      | Как находит колонку                                                                      |
|------------|------------------------------------------------------------------------------------------|
| `HEADER`   | Сканирует import-файл и ищет ячейку с текстом `headerName` (до 100 строк). **По умолчанию.** |
| `POSITION` | Берёт координаты маркера напрямую из шаблона, без поиска.                                |

```java
// Явно задать режим (только при использовании Excel-шаблона;
// JSON-шаблон определяет режим автоматически)
importParam.getParamsMap().put("for.user",
    new ImportInformation()
        .setClazz(User.class)
        .setBindingMode(BindingMode.POSITION));
```

---

## Как настроить импорт?

Импорт работает с двумя файлами: **шаблон** описывает расположение данных через маркеры,
**файл данных** содержит реальные значения в тех же позициях (или под теми же заголовками).

### Шаг 1 — Подготовьте класс данных

Класс должен иметь публичный конструктор без аргументов и setter-методы вида `setFieldName()`.
Имена полей класса должны совпадать с именами после точки в маркерах шаблона.

```java
// Маркер "test.company" → поле company → вызов setCompany(value)
public class Company {
    private String company;
    // getters / setters
}

// Маркер "for.user.salary" → поле salary → вызов setSalary(value)
public class User {
    private String name;
    private String age;
    private String salaryDate;
    private String salary;
    // getters / setters
}
```

### Шаг 2 — Создайте контейнер параметров

```java
var importParam = new ExcelImportParamCore();
```

### Шаг 3 — Зарегистрируйте классы по ключам

Ключ — это часть маркера **до последней точки**. Он должен совпадать с тем, что записано в шаблоне.

```java
// одиночный объект — маркеры вида "test.company"
importParam.getParamsMap().put("test",
    new ImportInformation().setClazz(Company.class));

// loop-список — маркеры вида "for.user.name", "for.user.salary"
importParam.getParamsMap().put("for.user",
    new ImportInformation().setClazz(User.class));
```

### Шаг 4 — Запустите импорт

**Через Excel-шаблон:**

```java
try (InputStream tmpl = new FileInputStream("template.xlsx");
     InputStream data = new FileInputStream("template_data.xlsx")) {
    ExcelImportUtil.importExcel(importParam, tmpl, data);
}
```

**Через JSON-шаблон:**

```java
try (InputStream tmpl = new FileInputStream("template.json");
     InputStream data = new FileInputStream("template_data.xlsx")) {
    ExcelImportUtil.importExcel(importParam, new JsonTemplateReader(tmpl), data);
}
```

### Шаг 5 — Получите результаты

```java
// одиночный объект — всегда один элемент в списке
Company company = (Company) importParam.getParamsMap().get("test").getLoopLst().get(0);

// loop-список — все прочитанные строки
List<Object> rows = importParam.getParamsMap().get("for.user").getLoopLst();
for (Object row : rows) {
    User user = (User) row;
    System.out.println(user.getName() + " — " + user.getSalary());
}
```

---

## Привязка к листам

Библиотека сопоставляет маркеры шаблона и данные **по имени листа**.
Лист в файле данных должен называться так же, как лист в шаблоне.

### Excel-шаблон

Имя листа определяется автоматически: библиотека сканирует все листы шаблона
и запоминает, на каком листе был найден каждый маркер.

```
template.xlsx:
  Лист "List1" → маркеры test.company, for.user.name, for.user.salary, ...
  Лист "List2" → маркеры for.user.name, for.user.salary, ...

template_data.xlsx:
  Лист "List1" → ячейки с реальными значениями
  Лист "List2" → строки с реальными данными
```

### JSON-шаблон

Имя листа задаётся явно в поле `sheetName` каждой записи:

```json
{
  "entries": [
    { "fieldName": "test.company",     "sheetName": "List1", "cellAddress": {"row": 3, "col": 0} },
    { "fieldName": "for.user.name",    "sheetName": "List1", "headerName": "NAME" },
    { "fieldName": "for.user.salary",  "sheetName": "List1", "headerName": "SALARY" }
  ]
}
```

> Если в файле данных нет листа с нужным именем — данные для этого листа не читаются,
> исключение не бросается.

### Несколько листов с одним ключом

Один ключ (например `for.user`) может встречаться на нескольких листах шаблона.
Результаты со всех листов **объединяются** в один список `getLoopLst()`.

```
template.xlsx:
  Лист "List1" → for.user.name, for.user.salary
  Лист "List2" → for.user.name, for.user.salary

// После импорта: getLoopLst() содержит строки с обоих листов
```

---

## Как выглядит шаблонный файл?

### Одиночный объект

| A            |
|--------------|
| test.company |

Данные читаются из **тех же координат** в import-файле.

### Loop-список

| A            | B           | C                | D           |
|--------------|-------------|------------------|-------------|
| NAME         | AGE         | SALARY_DAY       | SALARY      |
| for.user.name| for.user.age| for.user.salaryDate | for.user.salary |

- Строка **над** маркером — заголовок столбца (по нему находится колонка в import-файле).
- Все строки **под** заголовком в import-файле читаются в список.

---

## Какие типы данных поддерживаются?

| Тип ячейки Excel      | Поведение                                                              |
|-----------------------|------------------------------------------------------------------------|
| `NUMBER`              | Значение конвертируется в `String` через `BigDecimal.toPlainString()` |
| Любой другой          | Берётся `rawValue` ячейки как `String`                                |
| Дата (`m` в формате)  | Конвертируется в строку через `LocalDateTime.toString()`              |

Целевые поля объектов должны быть `String` или `Integer` (для служебного поля `row`).

---

## Как добавить обязательные поля?

```java
var info = new ImportInformation()
        .setClazz(User.class)
        .setRequiredFields(new ArrayList<>(List.of("name", "salary")));

importParam.getParamsMap().put("for.user", info);
```

При обнаружении заголовка в import-файле поле удаляется из списка `requiredFields`.
После импорта можно проверить, что список пуст:

```java
List<String> missing = importParam.getParamsMap().get("for.user").getRequiredFields();
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
public class Company {
    private String company;
    // getters / setters
}

public class User {
    private String name;
    private String age;
    private String salaryDate;
    private String salary;
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
    new ExportInformation().setDataList(List.of(company)));

// loop-список — ключ с префиксом "for."
exportParam.getParamsMap().put("for.user",
    new ExportInformation().setDataList(users));
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
Строка N−1:  | NAME  | AGE | SALARY_DAY | SALARY |   ← заголовки (из headerName)
Строка N:    | Alice | 30  | 2025-01-01 | 550    |   ← первый объект (на месте маркера)
Строка N+1:  | Bob   | 25  | 2025-02-01 | 430    |   ← второй объект
Строка N+2:  | Carol | 35  | 2025-03-01 | 670    |   ← третий объект
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
