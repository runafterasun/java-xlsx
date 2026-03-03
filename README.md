# excel-read-write — FAQ

## What is this?

A library for **importing and exporting** data from/to Excel files (`.xlsx`) using a template.
The template describes **where** the data is located (via markers in cells).

Under the hood it uses [fastexcel](https://github.com/dhatim/fastexcel) (reading and writing data)
and [Apache POI](https://poi.apache.org/) (streaming style reading from the template).

## How to add the dependency?

## Via Maven Central
Download from Maven Central <br>
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

## Via Maven Local

Publish the library to Maven Local:

```bash
./gradlew publishToMavenLocal
```

Then add it to your project:

**Gradle:**
```groovy
repositories {
    mavenLocal()
    mavenCentral()
}

dependencies {
    implementation 'ru.objectsfill:excel-read-write:0.0.1'
}
```

**Maven:**
```xml
<dependency>
    <groupId>ru.objectsfill</groupId>
    <artifactId>excel-read-write</artifactId>
    <version>0.0.1</version>
</dependency>
```

---

## How do markers work?

In the **template** Excel file, cells are marked with strings of the form:

```
<key>.<field>
```

| Marker              | Meaning                                                        |
|---------------------|----------------------------------------------------------------|
| `test.company`      | Single object with key `test`, field `company`                 |
| `for.user.salary`   | Loop list with key `for.user`, field `salary`                  |

- Keys without the `for.` prefix → **single object** (read once).
- Keys with the `for.` prefix → **list of objects** (all rows under the header are read).

---

## Template sources

Markers can be defined in two ways.

### Excel template

Create an `.xlsx` file with markers in the appropriate cells (`test.company`, `for.user.salary`, etc.).
The row **above** a loop marker is treated as the column header.

```java
// tmpl — template file with markers, data — file with actual data
ExcelImportUtil.importExcel(importParam, tmpl, data);
```

### JSON template

An alternative to the Excel template is a JSON file describing the same markers:

```json
{
  "entries": [
    { "fieldName": "test.company",     "sheetName": "List1", "cellAddress": {"row": 3, "col": 0} },
    { "fieldName": "for.user.salary",  "sheetName": "List1", "headerName": "SALARY" },
    { "fieldName": "for.user.age",     "sheetName": "List1", "headerName": "AGE" }
  ]
}
```

| Field         | Required | Description                                                    |
|---------------|----------|----------------------------------------------------------------|
| `fieldName`   | yes      | Marker in the form `key.field`                                 |
| `sheetName`   | yes      | Sheet name                                                     |
| `headerName`  | *        | Column header text (HEADER mode)                               |
| `cellAddress` | *        | Cell coordinates `{"row": N, "col": N}` (POSITION mode)       |

\* At least one of the two is required. If `headerName` is provided, HEADER mode is used.

```java
var reader = new JsonTemplateReader(new FileInputStream("template.json"));
ExcelImportUtil.importExcel(importParam, reader, data);
```

---

## Column binding modes (BindingMode)

Applies only to loop blocks (`for.` keys).

| Mode       | How the column is located                                                                     |
|------------|-----------------------------------------------------------------------------------------------|
| `HEADER`   | Scans the import file and looks for a cell with the text `headerName` (up to 100 rows). **Default.** |
| `POSITION` | Takes the marker coordinates directly from the template, without searching.                   |

```java
// Set the mode explicitly (only when using an Excel template;
// the JSON template determines the mode automatically)
importParam.getParamsMap().put("for.user",
    new ImportInformation()
        .setClazz(User.class)
        .setBindingMode(BindingMode.POSITION));
```

---

## How to set up import?

Import works with two files: the **template** describes the data layout via markers,
and the **data file** contains the actual values at the same positions (or under the same headers).

### Step 1 — Prepare the data class

The class must have a public no-argument constructor and setter methods of the form `setFieldName()`.
Field names must match the part after the dot in the template markers.

```java
// Marker "test.company" → field company → calls setCompany(value)
public class Company {
    private String company;
    // getters / setters
}

// Marker "for.user.salary" → field salary → calls setSalary(value)
public class User {
    private String name;
    private String age;
    private String salaryDate;
    private String salary;
    // getters / setters
}
```

### Step 2 — Create the parameter container

```java
var importParam = new ExcelImportParamCore();
```

### Step 3 — Register classes by key

The key is the part of the marker **up to the last dot**. It must match what is written in the template.

```java
// single object — markers like "test.company"
importParam.getParamsMap().put("test",
    new ImportInformation().setClazz(Company.class));

// loop list — markers like "for.user.name", "for.user.salary"
importParam.getParamsMap().put("for.user",
    new ImportInformation().setClazz(User.class));
```

### Step 4 — Run the import

**Using an Excel template:**

```java
try (InputStream tmpl = new FileInputStream("template.xlsx");
     InputStream data = new FileInputStream("template_data.xlsx")) {
    ExcelImportUtil.importExcel(importParam, tmpl, data);
}
```

**Using a JSON template:**

```java
try (InputStream tmpl = new FileInputStream("template.json");
     InputStream data = new FileInputStream("template_data.xlsx")) {
    ExcelImportUtil.importExcel(importParam, new JsonTemplateReader(tmpl), data);
}
```

### Step 5 — Get the results

```java
// single object — always one element in the list
Company company = (Company) importParam.getParamsMap().get("test").getLoopLst().get(0);

// loop list — all rows that were read
List<Object> rows = importParam.getParamsMap().get("for.user").getLoopLst();
for (Object row : rows) {
    User user = (User) row;
    System.out.println(user.getName() + " — " + user.getSalary());
}
```

---

## Sheet binding

The library matches template markers to data **by sheet name**.
The sheet in the data file must have the same name as the sheet in the template.

### Excel template

The sheet name is determined automatically: the library scans all sheets of the template
and remembers which sheet each marker was found on.

```
template.xlsx:
  Sheet "List1" → markers test.company, for.user.name, for.user.salary, ...
  Sheet "List2" → markers for.user.name, for.user.salary, ...

template_data.xlsx:
  Sheet "List1" → cells with actual values
  Sheet "List2" → rows with actual data
```

### JSON template

The sheet name is specified explicitly in the `sheetName` field of each entry:

```json
{
  "entries": [
    { "fieldName": "test.company",     "sheetName": "List1", "cellAddress": {"row": 3, "col": 0} },
    { "fieldName": "for.user.name",    "sheetName": "List1", "headerName": "NAME" },
    { "fieldName": "for.user.salary",  "sheetName": "List1", "headerName": "SALARY" }
  ]
}
```

> If the data file does not contain a sheet with the required name, data for that sheet is not read
> and no exception is thrown.

### Multiple sheets with the same key

One key (e.g. `for.user`) can appear on multiple template sheets.
Results from all sheets are **merged** into a single `getLoopLst()` list.

```
template.xlsx:
  Sheet "List1" → for.user.name, for.user.salary
  Sheet "List2" → for.user.name, for.user.salary

// After import: getLoopLst() contains rows from both sheets
```

---

## What does the template file look like?

### Single object

| A            |
|--------------|
| test.company |

Data is read from the **same coordinates** in the import file.

### Loop list

| A             | B            | C                   | D               |
|---------------|--------------|---------------------|-----------------|
| NAME          | AGE          | SALARY_DAY          | SALARY          |
| for.user.name | for.user.age | for.user.salaryDate | for.user.salary |

- The row **above** the marker is the column header (used to locate the column in the import file).
- All rows **below** the header in the import file are read into the list.

---

## What data types are supported?

| Excel cell type        | Behavior                                                               |
|------------------------|------------------------------------------------------------------------|
| `NUMBER`               | Value is converted to `String` via `BigDecimal.toPlainString()`        |
| Any other              | The cell's `rawValue` is used as a `String`                            |
| Date (`m` in format)   | Converted to a string via `LocalDateTime.toString()`                   |

Target object fields must be `String` or `Integer` (for the special `row` field).

---

## How to add required fields?

```java
var info = new ImportInformation()
        .setClazz(User.class)
        .setRequiredFields(new ArrayList<>(List.of("name", "salary")));

importParam.getParamsMap().put("for.user", info);
```

When a header is found in the import file, the field is removed from the `requiredFields` list.
After import you can verify the list is empty:

```java
List<String> missing = importParam.getParamsMap().get("for.user").getRequiredFields();
if (!missing.isEmpty()) {
    throw new IllegalStateException("Required fields not found: " + missing);
}
```

---

## Warnings

Non-critical issues (e.g. a field from the template is absent in the target class) do not throw an exception — they are collected in the warnings list:

```java
ExcelImportUtil.importExcel(importParam, tmpl, data);

List<String> warnings = importParam.getWarnings();
if (!warnings.isEmpty()) {
    warnings.forEach(System.out::println);
}
```

---

## How are errors handled?

Critical errors are wrapped in specialised exceptions with a context message:

| Exception               | When thrown                          |
|-------------------------|--------------------------------------|
| `ExcelImportException`  | Error during reading / import        |
| `ExcelExportException`  | Error during writing / export        |

Both are in the `ru.objectsfill.exception` package and extend `RuntimeException`.

```java
try (InputStream tmpl = ...; InputStream data = ...) {
    ExcelImportUtil.importExcel(importParam, tmpl, data);
} catch (ExcelImportException e) {
    log.error("Import error: {}", e.getMessage(), e);
}

try (InputStream tmpl = ...; OutputStream out = ...) {
    ExcelExportUtil.exportExcelWithStyles(exportParam, tmpl, out);
} catch (ExcelExportException e) {
    log.error("Export error: {}", e.getMessage(), e);
}
```

---

## How to set up export?

Export uses the same template as import: markers in cells indicate **where** to write the data.
Both template sources are supported — Excel file and JSON.

### Step 1 — Prepare the data classes

Classes must have a public no-argument constructor and getter methods of the form `getFieldName()`.
The same classes used for import work for export without any changes.

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

### Step 2 — Create the parameter container

```java
var exportParam = new ExcelExportParamCore();
```

### Step 3 — Register data by template key

Keys must match those used in the template (without the field name after the dot).

```java
// single object — key without the "for." prefix
exportParam.getParamsMap().put("test",
    new ExportInformation().setDataList(List.of(company)));

// loop list — key with the "for." prefix
exportParam.getParamsMap().put("for.user",
    new ExportInformation().setDataList(users));
```

> `setDataList` accepts `List<Object>`. For a single object pass a list with one element.

### Step 4 — Run the export

**Using an Excel template:**

```java
try (InputStream tmpl = new FileInputStream("template.xlsx");
     OutputStream out = new FileOutputStream("result.xlsx")) {
    ExcelExportUtil.exportExcel(exportParam, tmpl, out);
}
```

**Using a JSON template:**

```java
try (InputStream tmpl = new FileInputStream("template.json");
     OutputStream out = new FileOutputStream("result.xlsx")) {
    ExcelExportUtil.exportExcel(exportParam, new JsonTemplateReader(tmpl), out);
}
```

### Step 5 — Check warnings

Some situations do not interrupt the export but are recorded in the warnings list:
- a marker field is absent in the data class;
- no `cellAddress` is set for a marker (the field is skipped).

```java
List<String> warnings = exportParam.getWarnings();
if (!warnings.isEmpty()) {
    warnings.forEach(System.out::println);
}
```

---

### What the result looks like (Option A)

The marker row is replaced by the **first data row**. Subsequent rows are written below.

```
Row N−1:  | NAME  | AGE | SALARY_DAY | SALARY |   ← headers (from headerName)
Row N:    | Alice | 30  | 2025-01-01 | 550    |   ← first object (at the marker position)
Row N+1:  | Bob   | 25  | 2025-02-01 | 430    |   ← second object
Row N+2:  | Carol | 35  | 2025-03-01 | 670    |   ← third object
```

If `headerName` is not set (POSITION mode), the header row is not written.

---

### Export with styles

The `exportExcelWithStyles` method copies cell styles of single objects from the template to the result.
Apache POI reads the template **in streaming mode** (SAX), without loading the entire file into memory — only the first N rows are processed.

```java
// By default, styles are read from the first 100 rows of the template
try (InputStream tmpl = new FileInputStream("template.xlsx");
     OutputStream out = new FileOutputStream("result.xlsx")) {
    ExcelExportUtil.exportExcelWithStyles(exportParam, tmpl, out);
}

// Custom style-reading depth (if markers are below row 100)
try (InputStream tmpl = new FileInputStream("template.xlsx");
     OutputStream out = new FileOutputStream("result.xlsx")) {
    ExcelExportUtil.exportExcelWithStyles(exportParam, tmpl, out, 200);
}
```

| What is copied                  | Support |
|---------------------------------|---------|
| Bold, italic, underline         | ✓       |
| Font, size, font colour         | ✓       |
| Background colour (fill)        | ✓       |
| Borders (thin, medium, thick…)  | ✓       |
| Number format                   | ✓       |
| Horizontal alignment            | ✓       |
| Text wrapping                   | ✓       |
| Conditional formatting          | ✗       |
| Formulas                        | ✗       |

> Styles are applied only to **single-object** cells.
> Table rows (loop) are written without styles — by design.

---

### Export limitations

- When using `exportExcel` (without styles), the output file contains plain values — no template formatting.
- `exportExcelWithStyles` works only with an Excel template; a JSON template does not store styles.
- If `cellAddress` is not set for a marker, the field is skipped with a warning.
- When exporting with multiple sheets, loop-block data is written to **every** sheet that contains markers for that block in the template.

---

## Limitations

- No more than **100 rows** per sheet are read for single objects.
- Target class fields are populated via `setter` methods (format `setFieldName`).
- Target classes must have a **public no-argument constructor**.
- One loop block per sheet (multiple `for.` keys on the same sheet are not supported).
