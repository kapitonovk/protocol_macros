# Финализация сборки протокола — что делать

## 1. Правки в make_appx (3 строки, логика не тронута)

В `RunAppendixOnly` — путь к шаблону приложения:

```vba
' было
If Trim(CStr(numVal)) = "0.2" Then
' стало
If Trim(CStr(numVal)) = "0.4" Then
```

В `GetDocRefValue` — служебные таблицы:

```vba
' было
If UCase(conditionText) = "ДДУ" Then
    filePath = GetSourcePathByNumber(wbMain, "0.4")   ' -> ddu_data = 0.6
ElseIf UCase(conditionText) = "ДКП" Then
    filePath = GetSourcePathByNumber(wbMain, "0.3")   ' -> ПОДБОР писем = 0.5

' стало
If UCase(conditionText) = "ДДУ" Then
    filePath = GetSourcePathByNumber(wbMain, "0.6")
ElseIf UCase(conditionText) = "ДКП" Then
    filePath = GetSourcePathByNumber(wbMain, "0.5")
```

Больше в make_appx ничего не меняется. Всё, чего он не закрывает
(`{{I}}`, `{{КВАРТИРЫ_СЧЕТ}}`, блок 1.x.4 для не-ДДУ), дорабатывается
пост-обработкой в make_report.

## 2. Импорт модулей

- `make_body` — заменить целиком содержимым `make_body.txt`.
- `make_report` — новый модуль, содержимое `make_report.txt`.
- Точка входа: `BuildProtocol` (можно повесить на кнопку на листе Config).

Важно: в новом `make_body` удалены дубли `ColLetterToIndex`,
`FindHeaderColumn`, `FindSourceRowByNumber`, `GetDataScalarByPlaceholder`,
`ReplacePlaceholderInWord` — они остаются жить в make_appx. Без этого
вызов из третьего модуля даёт compile error «Ambiguous name detected».

## 3. Флаги в шапке make_report

```vba
Private Const FIX_HARDCODED_SECTION_NUMBERS As Boolean = True
Private Const DROP_DDU_BLOCK_FOR_NON_DDU   As Boolean = True
Private Const SHOW_RESULT                  As Boolean = True
```

- `FIX_HARDCODED_SECTION_NUMBERS` — заменяет в тексте жёстко прописанные
  «п. 1.1.1», «Приложению 1.1.2», «Приложении № 1.1.3» на номер текущего
  источника. Костыль на время; правильнее проставить в шаблонах `1.{{I}}.x`
  и выключить флаг.
- `DROP_DDU_BLOCK_FOR_NON_DDU` — для источников с Тип ≠ ДДУ вырезает из
  копии приложения блок «Приложение № 1.x.4» (рассрочка) вместе с Table 6,
  как и предполагает формула в Sources (столбец «Приложение 1.x.4»).

## 4. Как работает сборка

1. Читает Sources: 0.1 титул, 0.2 body, 0.3 лист согласования, 0.4 приложение.
2. Считает data-источники: строки, где «Номер» без точки (1, 2, 3 …),
   «Файл» заполнен и файл реально существует (`Dir`). Пустые строки 2…10
   просто игнорируются, количество источников любое.
3. Master-документ создаётся из титульника, в него подставляются
   `{{ПРОТ_НОМЕР}}`, `{{ДАТА1}}`, `{{ДАТА1ФОРМАТ}}`, `{{ДАТА2}}`.
4. Для каждого источника: data-файл открывается read-only →
   `BuildBodyDoc` (make_body) и `BuildSection_1_1_APPX` (make_appx) →
   пост-обработка номеров разделов → часть сохраняется во временный docx.
5. Лист согласования — отдельная часть.
6. Мердж в порядке: титул → body₁..N → лист согласования → приложение₁..N.
   Каждая часть вставляется через `Range.InsertFile` после разрыва раздела
   «со следующей страницы» — поэтому альбомные секции приложения не ломают
   портретный body.
7. Результат сохраняется рядом с main.xlsx:
   `protocol_<номер>_<ггггммдд_ччмм>.docx`, временные файлы удаляются,
   data-файлы закрываются без сохранения.

## 5. Мелочи, которые стоит поправить в шаблонах

- `template.appendix.docx`, абзац «Приложение № 1.{{I}.3» — не хватает
  закрывающей скобки. Код это уже подстраховывает (`1.{{I}.` → `1.N.`),
  но лучше починить в шаблоне.
- В body «в п. 1.1.1» и в приложении «Приложению 1.1.2 / № 1.1.3» —
  заменить на `1.{{I}}.x`.
- `FormatDateForProtocol` даёт «01» январь 2000 (Format "mmmm" в RU-локали
  выдаёт именительный падеж). Если нужен родительный — скажи, добавлю
  массив названий месяцев.
