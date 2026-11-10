# Спецификация: конвертация XLSX → TXT

Документ описывает полную логику конвертации XLSX → TXT из `converter.js` для точного воспроизведения на другой платформе. Все регулярные выражения, порядок операций, значения по умолчанию и краевые случаи соответствуют исходному коду.

---

## 1. Общая картина

Назначение: переводческая таблица XLSX (строки = переменные/ключи, колонки = языки) преобразуется в tab-separated TXT-файл с обратной ориентацией (строки = языки, колонки = переменные), с типографской очисткой текста и добавлением служебных колонок `code` и `Date`.

Пайплайн для каждого файла:

```
parseXLSX(file)                          → 2D-массив
extractCodeFromFilename(file.name)       → codeValue, hasCode
xlsxToTxt(data, options)                 → строка TXT + флаг foundNonBreakingHyphen
encodeString(txt, encoding) + BOM        → байты
скачивание файла <baseName>[-суффикс].txt
```

Поддерживается пакетная обработка: несколько XLSX-файлов обрабатываются последовательно в одном запуске, каждый даёт отдельный TXT и отдельную ссылку на скачивание. Ошибка одного файла не прерывает обработку остальных.

---

## 2. Входные параметры (настройки пользователя)

### 2.1. Кодировка вывода (`xlsxOutputEncoding`)

Одно из трёх значений: `utf-16le`, `macintosh`, `windows-1252`.

### 2.2. Флажки правил (объект `nbspSettings`)

| Ключ | По умолчанию | Что включает |
|---|---|---|
| `numericFormatting` | `true` | Числовое форматирование (мастер-переключатель для 3 подправил) |
| `numberSeparators` | `true` | Разделители тысяч/десятичных (подправило, без отдельного UI-флажка) |
| `percentFormatting` | `true` | Форматирование процентов (подправило, без отдельного UI-флажка) |
| `currencyFormatting` | `true` | Форматирование валют (подправило, без отдельного UI-флажка) |
| `unitSpacing` | `true` | NBSP между числом и единицей измерения |
| `unitCaseNormalization` | `true` | Нормализация регистра единиц (kw → kW) |
| `stripOuterQuotes` | `true` | Удаление внешних кавычек |
| `globalRules` | `true` | Глобальные NBSP-правила (инициалы, сокращения, даты) |
| `languageRules` | `true` | Языкоспецифичные правила |

Важно: у подправил `numberSeparators`, `percentFormatting`, `currencyFormatting` нет отдельных флажков в UI — они всегда `true` и применяются, если включён мастер-флажок `numericFormatting`.

### 2.3. Флажок суффикса кодировки (`ruleEncodingSuffix`)

Если включён, к имени выходного файла добавляется суффикс: `utf-16le` → `-Unicode`, `macintosh` → `-MacRoman`, `windows-1252` → `-Win1252`.

### 2.4. Производный параметр

`replaceNonBreakingHyphen = (encoding !== 'utf-16le')` — неразрывный дефис заменяется на обычный только при выводе в 8-битные кодировки.

---

## 3. Константы

```js
const NBSP = "\u00A0";
const NUMBER_PATTERN = '(?:\\d[\\d.,\\u00A0 ]*\\d|\\d)';
```

`NUMBER_PATTERN` — «число»: либо одна цифра, либо цифра + любые цифры/точки/запятые/NBSP/пробелы + цифра.

### 3.1. Единицы измерения (`UNIT_TOKENS`)

```js
['W', 'kW', 'MW', 'GW', 'hp', 'PS', 'CV',
 'V', 'kV', 'mV', 'A', 'mA', 'Hz', 'kHz', 'MHz', 'GHz', 'Ohm', 'Ω',
 'Wh', 'kWh', 'MWh', 'J', 'kJ', 'F', 'µF', 'μF', 'nF',
 'N', 'kN', 'Kgf', 'kgf', 'Pa', 'kPa', 'bar', 'Nm',
 'km', 'm', 'cm', 'mm', 'kg', 'g', 'mg', 'l', 'ml']
```

Составные единицы (буквы + опциональные цифры, разделённые `/` или `.`):

```js
const compositeUnits = '(?:[A-Za-zΩµμ]+[0-9]*(?:[./][A-Za-zΩµμ0-9]+)+)';
```

### 3.2. Известные языковые коды (`KNOWN_LANG_CODES`)

```js
new Set(['EN','FR','DE','ES','NL','IT','PT','RU','PL','CS',
         'SK','HU','RO','BG','HR','SL','SR','BS','MK','SQ',
         'EL','TR','DA','NO','SV','FI','ET','LV','LT','GA',
         'CY','MT','EU','CA','GL','JA','ZH','KO','AR','HE',
         'HI','TH','VI','ID','MS','TL','SW','UK','BE','US'])
```

### 3.3. Пары кавычек для `stripOuterQuotes`

```js
const pairs = { '"': '"', '«': '»', '\u201C': '\u201D', '\u201E': '\u201C' };
// то есть: "…" , «…» , “…” , „…“
```

---

## 4. Шаг 1. Парсинг XLSX

1. Файл читается как `ArrayBuffer`.
2. Библиотекой SheetJS: `XLSX.read(data, { type: 'array' })`.
3. Берётся **только первый лист**: `workbook.Sheets[workbook.SheetNames[0]]`.
4. Лист превращается в 2D-массив: `XLSX.utils.sheet_to_json(sheet, { header: 1 })`.
   - Пустые ячейки в середине строки → `undefined`; строки могут иметь разную длину («рваная» матрица).
   - Числовые ячейки приходят как числа (не строки) — дальнейший код везде делает `String(value)`.

Ожидаемая структура данных:

```
        колонка 0    колонка 1   колонка 2 ...
строка 0:  (пусто/метка)  en_EN      fr_FR      ← строка языковых кодов
строка 1:  Label1         Hello      Bonjour
строка 2:  Label2         World      Monde
```

---

## 5. Шаг 2. Извлечение кода из имени файла (`extractCodeFromFilename`)

1. Ищется первое вхождение `\[(.*?)\]` (текст в квадратных скобках). Если найдено и после `trim()` непусто → `{ value: <текст>, hasCode: true }`.
2. Иначе ищется `\((.*?)\)` (круглые скобки), та же логика.
3. Иначе → `{ value: 'not-set', hasCode: false }`.

Если `hasCode === false`, в лог добавляется уведомление: `No value in [] in filename. Using code = not-set.`

---

## 6. Шаг 3. Основное преобразование (`xlsxToTxt(data, options)`)

Вход: 2D-массив `data`, опции `{ replaceNonBreakingHyphen, codeValue, codeHasValue }`.

### 6.1. Проверка/добавление строки языковых кодов

Функция `isLanguageCode(str)` (после `trim()`):
- `/^[a-z]{2}_[A-Z]{2}$/` (например `en_EN`) → `true`;
- `/^[A-Za-z]{2}$/` и верхний регистр входит в `KNOWN_LANG_CODES` → `true`;
- иначе `false`.

`hasLanguageKeyRow(data)`: проверяет ячейки строки 0 **начиная с индекса 1** (индекс 0 — заголовок колонки ключей). Если хотя бы одна ячейка — языковой код, строка есть.

Если строки языков нет:
- вычисляется `maxCols` = максимальная длина строки по всему массиву (при пустом массиве — 2);
- в начало массива вставляется строка `['', 'en_EN', 'en_EN', ...]` (пустая ячейка + `en_EN` во всех остальных колонках до `maxCols`);
- в лог: `No language key row found in XLSX. Added default language "EN".`

### 6.2. Транспонирование (`transpose`)

Строки и колонки меняются местами. Количество колонок результата = максимальная длина исходной строки. Отсутствующие значения (`undefined`) заменяются на пустую строку `''`. После транспонирования: строки = языки, колонки = переменные.

### 6.3. Обработка каждой строки результата (индекс `index`)

Для каждой транспонированной строки, по порядку:

**а) Конвертация языкового кода (только `index > 0` и непустая ячейка 0):**
- `localeToShort`: если значение соответствует `/^([a-z]{2})_([A-Z]{2})$/i` → вторая группа в верхнем регистре (`en_EN` → `EN`); иначе значение без изменений;
- результат дополнительно приводится к верхнему регистру (`en` → `EN`);
- записывается в ячейку 0 и запоминается как `langCode`.

**б) Нормализация всех ячеек строки** (включая ячейку 0 с языковым кодом и всю строку заголовков `index === 0`) функцией `normalizeTxtCell` (см. раздел 7). Для каждой ячейки формируется ссылка на ячейку в координатах **исходного XLSX**:

```
originalCol = index + 1        // строка результата = колонка исходника
originalRow = cellIndex + 1
cellRef = columnToLetters(originalCol) + originalRow   // напр. "B3"
```

`columnToLetters` — стандартное преобразование 1 → A, 26 → Z, 27 → AA.

Одновременно проверяется наличие символа U+2011 (неразрывный дефис) в исходном значении ячейки — если найден хоть раз в файле, ставится флаг `foundNonBreakingHyphen`.

**в) Типографские правила** (только строки `index > 0` с непустой ячейкой 0): для ячеек с индекса 1 и далее вызывается `applyNbspRules(value, langCode)` (см. раздел 8), где `langCode` — уже нормализованный код из ячейки 0.

**г) Вставка колонки Date** на позицию 1 (после языкового кода):
- строка заголовков (`index === 0`): значение `'Date'`;
- остальные строки: текущая дата в формате `YY-MM-DD` (2-значный год, месяц и день с ведущими нулями).

**д) Вставка колонки code** на позицию 0:
- строка заголовков: значение `'code'`;
- остальные строки: `codeValue` (из имени файла либо `'not-set'`).

### 6.4. Сборка текста

```
txt = строки.map(row => row.join('\t')).join('\n')
```

Разделитель колонок — табуляция, разделитель строк — `\n` (без завершающего перевода строки).

Итоговая структура TXT:

```
code       (пусто)  Date      Label1   Label2
<code>     EN       26-07-12  Hello    World
<code>     FR       26-07-12  Bonjour  Monde
```

(вторая колонка в строке заголовков — то, что было в ячейке A1 исходника, обычно пусто)

---

## 7. Нормализация ячейки (`normalizeTxtCell(value, logger, context, langCode, options)`)

`null`/`undefined` → пустая строка. Иначе `String(value)` и по порядку:

1. **Неразрывный дефис** (`normalizeNonBreakingHyphen`, только если `options.replaceNonBreakingHyphen === true`): все `\u2011` → `-`. Лог-правило: `Non-breaking hyphen normalized`.
2. **Только для `langCode === 'FR'`** и включённом `languageRules` (`applyFrenchNarrowSpaces`):
   - все узкие неразрывные пробелы `\u202F` → NBSP `\u00A0` (лог: `FR narrow spaces normalized`);
   - удаление NBSP/узких пробелов внутри «ёлочек»: `/«[\u00A0\u202F]+/g` → `«` и `/[\u00A0\u202F]+»/g` → `»` (лог: `FR guillemet spacing removed`).
3. **Переносы строк**: `/[\r\n]+/g` → одиночный пробел. Лог (если были): `Line breaks removed`.
4. **Пробелы**: `/ {2,}/g` → одиночный пробел, затем `trim()`. Лог (если изменилось): `Whitespace normalization`.
5. **Внешние кавычки** (`stripOuterQuotes`, если включён `stripOuterQuotes`; строки короче 2 символов не трогаются):
   - если первый символ — открывающая кавычка из `pairs` и последний — соответствующая ей закрывающая → срезать оба (лог: `Outer quotes removed`);
   - если первый — открывающая, а последний — вообще не закрывающая кавычка → срезать только первый символ (лог: `Unmatched opening quote removed`);
   - если последний — закрывающая кавычка, а первый — не открывающая → срезать только последний (лог: `Unmatched closing quote removed`);
   - иначе без изменений.

---

## 8. Типографские правила (`applyNbspRules(value, langCode, options)`)

`null`/`undefined` → `''`. Пустая строка возвращается как есть. Правила применяются строго в следующем порядке (выход каждого шага — вход следующего):

```
1. applyNumericFormatting   (если numericFormatting)
   1a. applyNumberSeparators (если numberSeparators)
   1b. formatPercent         (если percentFormatting)
   1c. formatCurrency        (если currencyFormatting)
2. applyUnitSpacing          (если unitSpacing)
3. applyUnitCaseNormalization (если unitCaseNormalization)
4. applyGlobalRules          (если globalRules)
5. applyLanguageRules        (если languageRules)
```

`langCode` везде сравнивается в верхнем регистре.

### 8.1. Разделители чисел (`applyNumberSeparators`)

Определение разделителей по языку:
- разделитель тысяч: `FR`, `PL` → NBSP; `EN` → `,`; все остальные → `.`;
- десятичный разделитель: `EN` → `.`; все остальные → `,`.

Для каждого совпадения `NUMBER_PATTERN` в тексте:

**Парсинг числа (`parseNumberString`):**
1. NBSP заменяются обычными пробелами.
2. Подсчитывается количество `.` и `,` (`sepCount`); находится индекс последнего из них (`decimalIndex`).
3. Если разделитель есть: часть после него считается дробной, **кроме случая**, когда `sepCount === 1` и после разделителя ровно 3 цифры (тогда это разделитель тысяч, всё число — целая часть). Исключение из исключения: если цифры целой части (без разделителей/пробелов) равны `"0"` — дробная часть всё же признаётся (например `0.125`).
4. Из целой и дробной частей удаляются все `.`, `,` и пробельные символы. Если целая часть пуста — число не форматируется (возврат исходного совпадения).

**Форматирование целой части (`formatIntegerPart`):** если длина ≤ 4 символов — без изменений (т. е. `1000` остаётся `1000`); иначе группировка по 3 цифры справа выбранным разделителем тысяч: `/\B(?=(\d{3})+(?!\d))/g`.

Результат: `<целая часть>` либо `<целая часть><десятичный разделитель><дробная часть>`. Лог-правило: `Number separators`.

### 8.2. Проценты (`formatPercent`)

Регулярка: `(NUMBER_PATTERN)\s*%` (флаг `g`).
- `EN`: `число%` (без пробела);
- все остальные языки: `число<NBSP>%`.

Лог-правило: `Percent formatting`.

### 8.3. Валюты (`formatCurrency`)

Символы: `€`, `$`, `£`. Две регулярки (флаг `g`):
- prefix: `([€$£])\s*(NUMBER_PATTERN)`
- suffix: `(NUMBER_PATTERN)\s*([€$£])`

По языкам (порядок применения важен):
- `EN`: сначала suffix → `€100` (символ перед числом, без пробела), затем prefix → `€100`;
- `NL`: сначала suffix → `€<NBSP>100`, затем prefix → `€<NBSP>100` (символ перед числом с NBSP);
- `FR`, `DE`, `ES`, `PL`: сначала prefix → `100<NBSP>€`, затем suffix → `100<NBSP>€` (символ после числа с NBSP);
- прочие языки: без изменений.

Лог-правило: `Currency formatting`.

### 8.4. NBSP перед единицами (`applyUnitSpacing`)

```js
const simpleUnits = '(?:' + UNIT_TOKENS.join('|') + ')';
const spacedUnitsRegex = new RegExp('(' + NUMBER_PATTERN + ')\\s+(' + simpleUnits + '|' + compositeUnits + ')\\b', 'gi');
const tightUnitsRegex  = new RegExp('(' + NUMBER_PATTERN + ')('   + simpleUnits + '|' + compositeUnits + ')\\b', 'gi');
```

Сначала применяется `spacedUnitsRegex` (число, пробел(ы), единица), затем `tightUnitsRegex` (число вплотную к единице). Оба заменяют на `число<NBSP>единица`. Флаг `i` — регистронезависимо. Лог-правило: `NBSP before units`.

### 8.5. Нормализация регистра единиц (`applyUnitCaseNormalization`)

Ищется `(NUMBER_PATTERN)([\u00A0\s]+)(simpleUnits|compositeUnits)\b` (флаги `gi`) — число, затем пробел/NBSP, затем единица. Разделитель сохраняется как есть.

Таблица нормализации (`toLowerCase(единица)` → каноническая форма):

```
w→W  kw→kW  mw→MW  gw→GW  hp→hp  ps→PS  cv→CV
v→V  kv→kV  mv→mV  a→A  ma→mA  hz→Hz  khz→kHz  mhz→MHz  ghz→GHz  ohm→Ohm
wh→Wh  kwh→kWh  mwh→MWh  j→J  kj→kJ  f→F  uf→μF  µf→μF  μf→μF  nf→nF
n→N  kn→kN  kgf→kgf  pa→Pa  kpa→kPa  bar→bar
km→km  m→m  cm→cm  mm→mm  kg→kg  g→g  mg→mg  l→l  ml→ml  nm→Nm
```

- Если токен содержит `Ω` → результат `Ω`.
- Если единица не в таблице → без изменений.
- Составные единицы (содержат `/` или `.`): специальные случаи `kg/h/n` → `kg/h/N`, `m/s2` → `m/s²`; иначе каждый буквенно-цифровой сегмент (`/[A-Za-zΩµμ]+[0-9]*/g`) нормализуется по таблице отдельно (цифры на конце сегмента сохраняются).

Лог-правило: `Unit case normalization`.

### 8.6. Глобальные правила (`applyGlobalRules`)

Применяются по порядку, все с флагом `g` (первые — с учётом регистра, где указано иное — `gi`):

1. **Инициал + фамилия**:
   `/\b([A-ZÀ-ÖØ-ÞĄĆĘŁŃŚŹŻ])\.\s+([A-ZÀ-ÖØ-ÞĄĆĘŁŃŚŹŻ][A-Za-zÀ-ÖØ-öø-ÿĄĆĘŁŃŚŹŻąćęłńśźż'-]*)/g`
   → `Инициал.<NBSP>Фамилия`. Лог: `NBSP after initials`.
2. **Сокращения + число**: `/\b(p\.|№|Vol\.)\s+(\d+)/gi` → `сокр.<NBSP>число`. Лог: `NBSP after abbreviations`.
3. **Месяц + день**: `/\b(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+(\d{1,2},)/g` → `Месяц<NBSP>день,`. Лог: `NBSP in dates`.
4. **Год + слово year**: `/\b(\d{4})\s+(year)\b/gi` → `год<NBSP>year`. Лог: `NBSP in dates`.

### 8.7. Языкоспецифичные правила (`applyLanguageRules`)

Общий префикс для «коротких слов» — отрицательный lookbehind вместо `\b` (потому что `\b` в JS ломается на буквах с диакритикой):

```js
const LB = '(?<![A-Za-z\\xC0-\\xD6\\xD8-\\xF6\\xF8-\\xFF\\u0100-\\u024F])';
```

Правило «короткие слова»: `new RegExp(LB + '(слово1|слово2|...)\\s+', 'gi')` → `слово<NBSP>`. Списки по языкам:

| Язык | Слова | Лог-правило |
|---|---|---|
| EN | `a, an, the, I` | `EN NBSP after short words` |
| FR | `à, y, en, le, la, les, un, une, de, du` | `FR NBSP after short words` |
| ES | `a, e, i, o, u, y, la, el, un` | `ES NBSP after short words` |
| DE | `der, die, das, ein, eine, Dr., Prof., Hr., Fr.` (точки экранированы) | `DE NBSP after short words` |
| NL | `de, het, een, in, op, te, ten, ter` | `NL NBSP after short words` |
| PL | `a, e, i, o, u, w, z` | `PL NBSP after short words` |

Дополнительно **только для FR** (после правила коротких слов):
1. `/([^\s])\s*([;:!?])/g` → `символ<NBSP>знак` — NBSP перед `;` `:` `!` `?`. Лог: `FR NBSP before punctuation`.
2. `/«[\u00A0\u202F]+/g` → `«` и `/[\u00A0\u202F]+»/g` → `»` — удаление пробелов внутри «ёлочек». Лог: `FR guillemet spacing removed`.

Для остальных языков (не из таблицы) — без изменений.

---

## 9. Шаг 4. Кодирование и выгрузка

### 9.1. Имя выходного файла

```
baseName = имя файла без расширения .xlsx (регистронезависимо)
outputName = baseName [+ '-Unicode' | '-MacRoman' | '-Win1252' если включён суффикс]
файл: outputName + '.txt'
```

### 9.2. Кодирование (`encodeString(str, encoding)`)

- **`utf-16le`**: каждый JS-символ (UTF-16 code unit) → 2 байта little-endian. Перед содержимым добавляется BOM `FF FE`.
- **`macintosh` / `windows-1252`**: для каждого code unit по порядку:
  1. код < 0x80 → байт как есть;
  2. код есть в таблице соответствия (см. 9.3) → байт из таблицы;
  3. код в диапазоне 0x00A0–0x00FF → байт равен коду (passthrough Latin-1);
  4. иначе → байт `0x3F` (`?`).
- BOM для 8-битных кодировок не добавляется.

MIME выходного файла: `text/plain`.

### 9.3. Таблицы кодирования (Unicode codepoint → байт)

**Mac Roman** (`macRomanFromUnicode`) — полная таблица из кода:

```
00C4→80 00C5→81 00C7→82 00C9→83 00D1→84 00D6→85 00DC→86 00E1→87
00E0→88 00E2→89 00E4→8A 00E3→8B 00E5→8C 00E7→8D 00E9→8E 00E8→8F
00EA→90 00EB→91 00ED→92 00EC→93 00EE→94 00EF→95 00F1→96 00F3→97
00F2→98 00F4→99 00F6→9A 00F5→9B 00FA→9C 00F9→9D 00FB→9E 00FC→9F
2020→A0 00B0→A1 00A2→A2 00A3→A3 00A7→A4 2022→A5 00B6→A6 00DF→A7
00AE→A8 00A9→A9 2122→AA 00B4→AB 00A8→AC 2260→AD 00C6→AE 00D8→AF
221E→B0 00B1→B1 2264→B2 2265→B3 00A5→B4 00B5→B5 2202→B6 2211→B7
220F→B8 03C0→B9 222B→BA 00AA→BB 00BA→BC 03A9→BD 00E6→BE 00F8→BF
00BF→C0 00A1→C1 00AC→C2 221A→C3 0192→C4 2248→C5 2206→C6 00AB→C7
00BB→C8 2026→C9 00A0→CA 00C0→CB 00C3→CC 00D5→CD 0152→CE 0153→CF
2013→D0 2014→D1 201C→D2 201D→D3 2018→D4 2019→D5 00F7→D6 25CA→D7
00FF→D8 0178→D9 2044→DA 20AC→DB 2039→DC 203A→DD FB01→DE FB02→DF
2021→E0 00B7→E1 201A→E2 201E→E3 2030→E4 00C2→E5 00CA→E6 00C1→E7
00CB→E8 00C8→E9 00CD→EA 00CE→EB 00CF→EC 00CC→ED 00D3→EE 00D4→EF
F8FF→F0 00D2→F1 00DA→F2 00DB→F3 00D9→F4 0131→F5 02C6→F6 02DC→F7
00AF→F8 02D8→F9 02D9→FA 02DA→FB 00B8→FC 02DD→FD 02DB→FE 02C7→FF
```

**Windows-1252** (`win1252FromUnicode`):

```
20AC→80 201A→82 0192→83 201E→84 2026→85 2020→86 2021→87 02C6→88
2030→89 0160→8A 2039→8B 0152→8C 017D→8E 2018→91 2019→92 201C→93
201D→94 2022→95 2013→96 2014→97 02DC→98 2122→99 0161→9A 203A→9B
0153→9C 017E→9E 0178→9F
```

(для Win1252 диапазон 0xA0–0xFF покрывается passthrough-правилом п. 3 из 9.2; символы не из таблицы и не из Latin-1 → `?`)

### 9.4. Уведомление о неразрывном дефисе

Если в файле был найден U+2011 (`foundNonBreakingHyphen === true`) и выбрана кодировка не `utf-16le`, в лог добавляется уведомление:
`Non-breaking hyphen found. Use Unicode output to preserve it. Otherwise the character will be replaced with a regular hyphen, which is safe.`

---

## 10. Логирование

Каждая запись лога: `{ rule, before, after, context, filename }`, где `context = { cellRef, lang }` (может быть `null` для уведомлений).

Типы записей:
- обычное изменение: правило + значение до + значение после + ссылка на ячейку;
- `Notice` — информационное сообщение (текст в `before`);
- `Error` — ошибка обработки файла (текст в `before`).

Правила:
- запись добавляется только если `before !== after` (для правил замены);
- счётчик в заголовке лога считает все записи, кроме `Notice` (текст `N changes`);
- в UI для правил `Whitespace normalization` и `Line breaks removed` показывается только значение «до», с визуализацией удаляемых пробелов/переносов символом `□`;
- `cellRef` отображается в формате «буква колонки + номер строки» исходного XLSX (например `B3`), номер строки выделяется жирным.

Список полных названий лог-правил (точные строки, используются как идентификаторы):

```
Non-breaking hyphen normalized
FR narrow spaces normalized
FR guillemet spacing removed
Line breaks removed
Whitespace normalization
Outer quotes removed
Unmatched opening quote removed
Unmatched closing quote removed
Number separators
Percent formatting
Currency formatting
NBSP before units
Unit case normalization
NBSP after initials
NBSP after abbreviations
NBSP in dates
EN NBSP after short words
FR NBSP after short words
ES NBSP after short words
DE NBSP after short words
NL NBSP after short words
PL NBSP after short words
FR NBSP before punctuation
Notice
Error
```

---

## 11. Обработка ошибок и пакетный режим

- Перед запуском конвертации очищаются лог и область ссылок на скачивание; заголовок лога: `XLSX → TXT log`.
- Файлы обрабатываются последовательно; для каждого файла `currentLogFilename = file.name` (имя добавляется к каждой записи лога).
- Исключение при обработке файла → сообщение `<имя файла>: <текст ошибки>` добавляется в список ошибок и в лог как `Error`; обработка продолжается со следующего файла.
- После обработки всех файлов: рендер лога; если были ошибки — вывод `Errors: <ошибка1>; <ошибка2>; ...`.
- При добавлении файлов в список дубликаты отсеиваются по паре «имя + размер»; принимаются только файлы с расширением `.xlsx` (регистронезависимо).

---

## 12. Карта соответствия функций исходника

| Функция в `converter.js` | Раздел спецификации |
|---|---|
| `convertXlsxToTxt` | 1, 11 |
| `parseXLSX` | 4 |
| `extractCodeFromFilename` | 5 |
| `xlsxToTxt` | 6 |
| `hasLanguageKeyRow`, `isLanguageCode` | 6.1 |
| `transpose` | 6.2 |
| `localeToShort` | 6.3а |
| `normalizeTxtCell` | 7 |
| `normalizeNonBreakingHyphen` | 7 (п. 1) |
| `applyFrenchNarrowSpaces` | 7 (п. 2) |
| `stripOuterQuotes` | 7 (п. 5) |
| `applyNbspRules` | 8 |
| `applyNumberSeparators` | 8.1 |
| `formatPercent` | 8.2 |
| `formatCurrency` | 8.3 |
| `applyUnitSpacing` | 8.4 |
| `applyUnitCaseNormalization` | 8.5 |
| `applyGlobalRules` | 8.6 |
| `applyLanguageRules` | 8.7 |
| `getCurrentDate` | 6.3г |
| `columnToLetters` | 6.3б |
| `encodeString`, `encodeUTF16LE`, `downloadTXT` | 9 |
| `macRomanFromUnicode`, `win1252FromUnicode` | 9.3 |
| `addLog`, `addLogNotice`, `addLogError`, `renderLog` | 10 |

---

## 13. Важные тонкости для точного воспроизведения

1. **Порядок правил критичен**: числа → проценты → валюты → NBSP перед единицами → регистр единиц → глобальные → языковые. Нормализация ячейки (`normalizeTxtCell`) выполняется **до** типографских правил.
2. `normalizeTxtCell` применяется ко **всем** ячейкам, включая строку заголовков и колонку языковых кодов; `applyNbspRules` — только к контентным ячейкам (строки > 0, колонки > 0 до вставки `code`/`Date`).
3. На момент вызова `normalizeTxtCell` для строк данных `langCode` уже нормализован (важно для FR-правил внутри нормализации).
4. Целые числа длиной 4 цифры (например `1000`) разделителями тысяч **не** форматируются; от 5 цифр — форматируются.
5. `1.500` (один разделитель + ровно 3 цифры после) трактуется как `1500` (разделитель тысяч), но `0.125` — как дробь.
6. Колонки `code` и `Date` добавляются **после** всех текстовых преобразований и не подвергаются им.
7. `cellRef` в логе указывает на координаты **исходного XLSX** (до транспонирования и вставки колонок).
8. Регулярки единиц регистронезависимы (`gi`), поэтому `100KW` тоже матчится; нормализация регистра затем приводит к `kW`.
9. Замена U+2011 зависит от выбранной кодировки вывода и происходит в самом начале обработки ячейки.
10. Лог фиксирует изменение только при фактической разнице «до/после».
