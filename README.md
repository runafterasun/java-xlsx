# excel-read-write — FAQ

## Что это?

Библиотека для импорта данных из Excel-файлов (`.xlsx`) по шаблону.
Шаблон описывает **где** искать данные (через маркеры в ячейках), а import-файл содержит **сами данные**.

Под капотом используется [fastexcel-reader](https://github.com/dhatim/fastexcel).

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

## Как обрабатываются ошибки?

Все ошибки оборачиваются в `ExcelImportException` (`ru.objectsfill.exception`) с контекстным сообщением.

```java
try (InputStream tmpl = ...; InputStream data = ...) {
    ExcelImportUtil.importExcel(importParam, tmpl, data);
} catch (ExcelImportException e) {
    log.error("Ошибка импорта: {}", e.getMessage(), e);
}
```

---

## Ограничения

- Читается не более **100 строк** на лист для одиночных объектов.
- Поля целевых классов заполняются через `setter`-методы (формат `setFieldName`).
- Целевые классы должны иметь **публичный конструктор без аргументов**.
- Один loop-блок на лист (несколько `for.`-ключей на одном листе не поддерживается).
