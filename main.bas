Attribute VB_Name = "make_section_1_1_unified"
Option Explicit

' ============================================================
' ЕДИНАЯ ТОЧКА ВХОДА
' Можно запускать:
'   - только BODY
'   - только APPX
'   - BODY + APPX вместе
' ============================================================

Sub RunSection_1_1_All()

    Call RunSection_1_1(True, True)

End Sub

Sub RunSection_1_1_BodyOnly()

    Call RunSection_1_1(True, False)

End Sub

Sub RunSection_1_1_AppendixOnly()

    Call RunSection_1_1(False, True)

End Sub

Private Sub RunSection_1_1(buildBody As Boolean, buildAppendix As Boolean)

    Dim wbMain As Workbook
    Dim wbData As Workbook

    Dim wdApp As Object
    Dim wdBody As Object
    Dim wdAppx As Object

    Dim bodyPath As String
    Dim appendixPath As String
    Dim dataPath As String

    Dim sourceRowSection1 As Long

    Set wbMain = ThisWorkbook

    ' Пути берём через заголовки Sources, а не по "жёстким" номерам колонок
    bodyPath = GetSourcePathByNumber(wbMain, "0.1")
    appendixPath = GetSourcePathByNumber(wbMain, "0.2")
    dataPath = GetSourcePathByNumber(wbMain, "1")

    sourceRowSection1 = FindSourceRowByNumber(wbMain, "1")

    If buildBody Then
        If bodyPath = "" Then
            MsgBox "Не найден путь к BODY-шаблону (Номер = 0.1) в Sources."
            Exit Sub
        End If
    End If

    If buildAppendix Then
        If appendixPath = "" Then
            MsgBox "Не найден путь к APPX-шаблону (Номер = 0.2) в Sources."
            Exit Sub
        End If
    End If

    If dataPath = "" Then
        MsgBox "Не найден путь к data-файлу (Номер = 1) в Sources."
        Exit Sub
    End If

    If sourceRowSection1 = 0 Then
        MsgBox "Не найдена строка в Sources для Номер = 1."
        Exit Sub
    End If

    Set wbData = Workbooks.Open(Filename:=dataPath, ReadOnly:=True, UpdateLinks:=False)

    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = True

    If buildBody Then
        Set wdBody = wdApp.Documents.Add(Template:=bodyPath)
        Call BuildSection_1_1_BODY_MAIN(wbMain, wbData, wdBody, sourceRowSection1)
    End If

    If buildAppendix Then
        Set wdAppx = wdApp.Documents.Add(Template:=appendixPath)
        Call BuildSection_1_1_APPX(wbMain, wbData, wdAppx, sourceRowSection1)
    End If

    ' Сохранение отключено намеренно — включить при необходимости:
    ' If Not wdBody Is Nothing Then wdBody.SaveAs2 Filename:=wbMain.Path & "\test_body_result.docx"
    ' If Not wdAppx Is Nothing Then wdAppx.SaveAs2 Filename:=wbMain.Path & "\test_appendix_result.docx"

    ' wbData.Close SaveChanges:=False

End Sub


' ============================================================
' ОБЩИЕ УТИЛИТЫ
' ============================================================

Function ColLetterToIndex(colLetter As String) As Long
    colLetter = UCase(Trim(colLetter))
    ColLetterToIndex = Range(colLetter & "1").Column
End Function


Function FindHeaderColumn(ws As Worksheet, headerText As String) As Long

    Dim lastCol As Long
    Dim c As Long

    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    For c = 1 To lastCol
        If Trim(CStr(ws.Cells(1, c).Value)) = Trim(headerText) Then
            FindHeaderColumn = c
            Exit Function
        End If
    Next c

    FindHeaderColumn = 0

End Function


Function FindSourceRowByNumber(wbMain As Workbook, sourceNumber As Variant) As Long

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim numCol As Long
    Dim r As Long

    Set ws = wbMain.Worksheets("Sources")
    numCol = FindHeaderColumn(ws, "Номер")

    If numCol = 0 Then
        FindSourceRowByNumber = 0
        Exit Function
    End If

    lastRow = ws.Cells(ws.Rows.Count, numCol).End(xlUp).Row

    For r = 2 To lastRow
        If Trim(CStr(ws.Cells(r, numCol).Value)) = Trim(CStr(sourceNumber)) Then
            FindSourceRowByNumber = r
            Exit Function
        End If
    Next r

    FindSourceRowByNumber = 0

End Function


Function GetSourceValueByHeader(wbMain As Workbook, sourceRow As Long, headerText As String) As String

    Dim ws As Worksheet
    Dim colIdx As Long

    Set ws = wbMain.Worksheets("Sources")
    colIdx = FindHeaderColumn(ws, headerText)

    If colIdx = 0 Then
        GetSourceValueByHeader = ""
    Else
        GetSourceValueByHeader = Trim(CStr(ws.Cells(sourceRow, colIdx).Value))
    End If

End Function


Function GetSourcesValue(wbMain As Workbook, sourceRow As Long, columnHeader As String) As String

    Dim ws As Worksheet
    Dim colNum As Long

    Set ws = wbMain.Worksheets("Sources")
    colNum = FindHeaderColumn(ws, columnHeader)

    If colNum = 0 Then
        GetSourcesValue = ""
        Exit Function
    End If

    GetSourcesValue = CStr(ws.Cells(sourceRow, colNum).Value)

End Function


Function GetSourcePathByNumber(wbMain As Workbook, sourceNumber As String) As String

    Dim ws As Worksheet
    Dim sourceRow As Long
    Dim colPath As Long

    Set ws = wbMain.Worksheets("Sources")
    sourceRow = FindSourceRowByNumber(wbMain, sourceNumber)
    colPath = FindHeaderColumn(ws, "Путь")

    If sourceRow = 0 Or colPath = 0 Then
        GetSourcePathByNumber = ""
    Else
        GetSourcePathByNumber = Trim(CStr(ws.Cells(sourceRow, colPath).Value))
    End If

End Function


Function GetConfigValue(wbMain As Workbook, placeholderText As String) As String

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    Set ws = wbMain.Worksheets("Config")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow
        If Trim(CStr(ws.Cells(r, 2).Value)) = Trim(placeholderText) Then
            GetConfigValue = CStr(ws.Cells(r, 3).Value)
            Exit Function
        End If
    Next r

    GetConfigValue = ""

End Function


Function GetDataScalarByPlaceholder(wbMain As Workbook, wbData As Workbook, placeholderText As String) As String

    Dim wsMapper As Worksheet
    Dim lastRow As Long
    Dim r As Long

    Dim sheetName As String
    Dim cellRef As String
    Dim wsData As Worksheet

    Dim phCol As Long
    Dim sheetCol As Long
    Dim refCol As Long

    Set wsMapper = wbMain.Worksheets("Mapper")

    phCol = FindHeaderColumn(wsMapper, "Placeholder")
    sheetCol = FindHeaderColumn(wsMapper, "ExcelSheet")
    refCol = FindHeaderColumn(wsMapper, "ExcelColumnRef")

    If phCol = 0 Or sheetCol = 0 Or refCol = 0 Then
        GetDataScalarByPlaceholder = ""
        Exit Function
    End If

    lastRow = wsMapper.Cells(wsMapper.Rows.Count, phCol).End(xlUp).Row

    For r = 2 To lastRow
        If Trim(CStr(wsMapper.Cells(r, phCol).Value)) = Trim(placeholderText) Then

            sheetName = Trim(CStr(wsMapper.Cells(r, sheetCol).Value))
            cellRef = Trim(CStr(wsMapper.Cells(r, refCol).Value))

            If sheetName <> "" And cellRef <> "" Then
                Set wsData = wbData.Worksheets(sheetName)
                GetDataScalarByPlaceholder = CStr(wsData.Range(cellRef).Value)
                Exit Function
            End If
        End If
    Next r

    GetDataScalarByPlaceholder = ""

End Function


Sub ReplacePlaceholderInWord(wdDoc As Object, findText As String, replaceText As String)

    With wdDoc.Content.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = findText
        .Replacement.Text = replaceText
        .Forward = True
        .Wrap = 1
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .Execute Replace:=2
    End With

End Sub


Function FormatDateForProtocol(dateText As String) As String

    Dim d As Date
    Dim monthName As String

    If IsDate(dateText) Then
        d = CDate(dateText)
        monthName = LCase(Format(d, "mmmm"))
        FormatDateForProtocol = "«" & Format(d, "dd") & "» " & monthName & " " & Format(d, "yyyy")
    Else
        FormatDateForProtocol = dateText
    End If

End Function


Function CleanNumericCell(rawText As String) As String

    Dim t As String

    t = rawText
    t = Replace(t, Chr(160), " ")
    t = Replace(t, vbCr, "")
    t = Replace(t, vbLf, "")
    t = Trim(t)

    CleanNumericCell = t

End Function


Sub SafeSetCellText(wdTable As Object, rowIndex As Long, colIndex As Long, txt As String)
    On Error GoTo SkipCell
    wdTable.Cell(rowIndex, colIndex).Range.Text = CleanNumericCell(txt)
    Exit Sub
SkipCell:
End Sub


Function SafeGetCellText(wdTable As Object, rowIndex As Long, colIndex As Long) As String
    On Error GoTo SkipCell
    SafeGetCellText = wdTable.Cell(rowIndex, colIndex).Range.Text
    Exit Function
SkipCell:
    SafeGetCellText = ""
End Function


Function SafeGetRowText(wdTable As Object, rowIndex As Long) As String
    On Error GoTo SkipRow
    SafeGetRowText = wdTable.Rows(rowIndex).Range.Text
    SafeGetRowText = Replace(SafeGetRowText, Chr(13), "")
    SafeGetRowText = Replace(SafeGetRowText, Chr(7), "")
    Exit Function
SkipRow:
    SafeGetRowText = ""
End Function


Function ExtractCategoryCode(rawText As String) As String

    Dim t As String
    Dim p1 As Long
    Dim p2 As Long

    t = Replace(rawText, Chr(13), "")
    t = Replace(t, Chr(7), "")

    p1 = InStr(1, t, "{{")
    p2 = InStr(1, t, "}}")

    If p1 > 0 And p2 > p1 Then
        ExtractCategoryCode = Mid$(t, p1 + 2, p2 - p1 - 2)
    Else
        ExtractCategoryCode = ""
    End If

End Function


Function InsertEmptyRowsAfter(wdTable As Object, afterRowIndex As Long, rowsToAdd As Long) As Long

    Dim i As Long
    Dim insertBeforeIndex As Long

    InsertEmptyRowsAfter = afterRowIndex + 1

    For i = 1 To rowsToAdd
        insertBeforeIndex = afterRowIndex + i

        If insertBeforeIndex > wdTable.Rows.Count Then
            wdTable.Rows.Add
        Else
            wdTable.Rows.Add BeforeRow:=wdTable.Rows(insertBeforeIndex)
        End If
    Next i

End Function


Function SumColumnByLetter(wsData As Worksheet, colLetter As String, startRow As Long, lastRow As Long) As Double

    Dim r As Long
    Dim colIdx As Long
    Dim v As Variant
    Dim total As Double

    colIdx = ColLetterToIndex(colLetter)
    total = 0

    For r = startRow To lastRow
        v = wsData.Cells(r, colIdx).Value
        If IsNumeric(v) Then
            total = total + CDbl(v)
        End If
    Next r

    SumColumnByLetter = total

End Function


' ============================================================
' BODY — ОРКЕСТРАТОР
' ============================================================
Sub BuildSection_1_1_BODY_MAIN(wbMain As Workbook, wbData As Workbook, wdBody As Object, sourceRow As Long)

    Call FillTable_BODY_MAIN(wbMain, wbData, wdBody)
    Call ApplyScalarsToDocument(wbMain, wbData, wdBody, sourceRow)

End Sub


Sub FillTable_BODY_MAIN(wbMain As Workbook, wbData As Workbook, wdBody As Object)

    Dim wsMapper As Worksheet
    Dim wsData As Worksheet
    Dim wdTable As Object

    Dim firstRow As Long
    Dim lastMapRow As Long
    Dim r As Long

    Dim excelSheetName As String
    Dim keyColLetter As String
    Dim startRow As Long
    Dim keyColIdx As Long
    Dim lastDataRow As Long
    Dim rowCount As Long

    Dim i As Long
    Dim wordRow As Long

    Set wsMapper = wbMain.Worksheets("Mapper")

    ' Найти первую строку для BODY_MAIN в Mapper
    lastMapRow = wsMapper.Cells(wsMapper.Rows.Count, 2).End(xlUp).Row
    firstRow = 0

    For r = 2 To lastMapRow
        If Trim(CStr(wsMapper.Cells(r, 2).Value)) = "BODY_MAIN" Then
            firstRow = r
            excelSheetName = Trim(CStr(wsMapper.Cells(r, 5).Value))   ' ExcelSheet
            keyColLetter = Trim(CStr(wsMapper.Cells(r, 6).Value))     ' ExcelColumnRef (буква)
            startRow = CLng(wsMapper.Cells(r, 7).Value)               ' StartRow
            Exit For
        End If
    Next r

    If firstRow = 0 Then
        MsgBox "В Mapper нет строк для BODY_MAIN."
        Exit Sub
    End If

    ' Определяем длину данных по ключевой колонке
    Set wsData = wbData.Worksheets(excelSheetName)
    keyColIdx = ColLetterToIndex(keyColLetter)

    lastDataRow = wsData.Cells(wsData.Rows.Count, keyColIdx).End(xlUp).Row
    If lastDataRow < startRow Then
        MsgBox "В листе " & excelSheetName & " нет данных для BODY_MAIN."
        Exit Sub
    End If

    rowCount = lastDataRow - startRow + 1

    ' Берём таблицу квартир
    Set wdTable = wdBody.Tables(2)

    ' Чистим старые строки данных (оставляем шапку + строку-образец)
    Do While wdTable.Rows.Count > 2
        wdTable.Rows(3).Delete
    Loop

    ' Убедимся, что строка-образец (2) существует
    If wdTable.Rows.Count < 2 Then
        wdTable.Rows.Add
    End If

    ' Добавляем строки под все данные:
    ' первая строка данных пойдёт в уже существующую строку 2,
    ' поэтому добавляем пустые строки ниже неё
    If rowCount > 1 Then
        Call InsertEmptyRowsAfter(wdTable, 2, rowCount - 1)
    End If

    ' Заполняем таблицу
    ' В Word: строка 1 – шапка, с 2 – данные
    For i = 0 To rowCount - 1
        wordRow = 2 + i

        ' Остальные колонки по Mapper’у (BODY_MAIN)
        For r = 2 To lastMapRow
            If Trim(CStr(wsMapper.Cells(r, 2).Value)) = "BODY_MAIN" Then

                Dim colOrder As Long
                Dim colLetter As String
                Dim sheetName As String
                Dim valueRow As Long
                Dim colIdx As Long

                colOrder = CLng(wsMapper.Cells(r, 3).Value)
                sheetName = Trim(CStr(wsMapper.Cells(r, 5).Value))
                colLetter = Trim(CStr(wsMapper.Cells(r, 6).Value))

                Set wsData = wbData.Worksheets(sheetName)
                colIdx = ColLetterToIndex(colLetter)
                valueRow = startRow + i

                Call SafeSetCellText(wdTable, wordRow, colOrder, wsData.Cells(valueRow, colIdx).Text)
            End If
        Next r
    Next i

End Sub


Sub ApplyScalarsToDocument(wbMain As Workbook, wbData As Workbook, wdDoc As Object, sourceRow As Variant)

    Dim protNum As String
    Dim date1 As String
    Dim date1Formatted As String
    Dim date2 As String
    Dim appendixList As String
    Dim addressVal As String

    If CLng(sourceRow) = 0 Then
        MsgBox "Не найдена строка в Sources для раздела 1."
        Exit Sub
    End If

    protNum = GetConfigValue(wbMain, "{{ПРОТ_НОМЕР}}")
    date1 = GetConfigValue(wbMain, "{{ДАТА1}}")
    date1Formatted = FormatDateForProtocol(date1)

    ' Пока временная логика:
    ' если отдельного правила нет, дата2 = дата1 + 90 дней
    If IsDate(date1) Then
        date2 = Format(DateAdd("d", 90, CDate(date1)), "dd.mm.yyyy")
    Else
        date2 = date1
    End If

    appendixList = GetSourcesValue(wbMain, CLng(sourceRow), "Список приложений")
    addressVal = GetDataScalarByPlaceholder(wbMain, wbData, "{{АДРЕС}}")

    Call ReplacePlaceholderInWord(wdDoc, "{{ПРОТ_НОМЕР}}", protNum)
    Call ReplacePlaceholderInWord(wdDoc, "{{ДАТА1}}", date1)
    Call ReplacePlaceholderInWord(wdDoc, "{{ДАТА1ФОРМАТ}}", date1Formatted)
    Call ReplacePlaceholderInWord(wdDoc, "{{ДАТА2}}", date2)
    Call ReplacePlaceholderInWord(wdDoc, "{{ПРИЛОЖЕНИЯ}}", appendixList)
    Call ReplacePlaceholderInWord(wdDoc, "{{АДРЕС}}", addressVal)

End Sub


' ============================================================
' APPX — ОРКЕСТРАТОР
' ============================================================
Sub BuildSection_1_1_APPX(wbMain As Workbook, wbData As Workbook, wdAppx As Object, sourceRow As Long)

    Dim activeCats As Object

    ' Table 1 — шапка (округ/район/адрес), простые скаляры
    Call FillScalars_APPX_111a(wbMain, wbData, wdAppx)

    ' Вариативные абзацы
    Call FillConditionalPhrases_APPX(wbMain, wbData, wdAppx, sourceRow)

    ' Считаем активные категории один раз — используются в Table 2 и в подвале Table 3
    Set activeCats = GetActiveCategories(wbMain, wbData)

    ' Table 2 — список категорий с итоговыми суммами
    Call FillTable_APPX_111b(wdAppx, activeCats)

    ' Table 3 — аналоги-аукционы, построчно + подвал по категориям
    Call FillTable_APPX_112(wbMain, wbData, wdAppx, activeCats)

    ' Table 4 — матрица корректировки (коэффициенты)
    Call FillTable_APPX_113a(wbMain, wbData, wdAppx)

    ' Table 5 — ценообразование по квартирам + 3 суммы в подвале
    Call FillTable_APPX_113b(wbMain, wbData, wdAppx)

    ' Table 6 — варианты первоначального взноса
    Call FillTable_APPX_114(wbMain, wbData, wdAppx)

End Sub


' ============================================================
' УСЛОВНЫЕ ПЛЕЙСХОЛДЕРЫ ДЛЯ ПРИЛОЖЕНИЯ
' {{ОБЪЕКТ_ШАБЛОННАЯ_ФРАЗА}}
' {{СОГЛАСОВАНИЕ_ИСТОЧНИКИ}}
' ============================================================

Sub FillConditionalPhrases_APPX(wbMain As Workbook, wbData As Workbook, wdDoc As Object, sourceRow As Long)

    Dim dealType As String
    Dim developerType As String
    Dim objectAddress As String

    Dim objectCondition As String
    Dim objectPhrase As String

    Dim approvalTemplate As String
    Dim approvalInnerPlaceholder As String
    Dim approvalData As String
    Dim approvalPhrase As String

    dealType = Trim(CStr(GetSourceValueByHeader(wbMain, sourceRow, "Тип")))
    developerType = Trim(CStr(GetSourceValueByHeader(wbMain, sourceRow, "Застройщик (фонд / не фонд)")))
    objectAddress = Trim(GetDataScalarByPlaceholder(wbMain, wbData, "{{АДРЕС}}"))

    ' ---- 1. {{ОБЪЕКТ_ШАБЛОННАЯ_ФРАЗА}} ----
    If UCase(dealType) = "ДКП" And LCase(developerType) = "фонд" Then
        objectCondition = "ДКП_фонд"
    Else
        objectCondition = "Иное"
    End If

    objectPhrase = GetPhraseValue(wbMain, "{{ОБЪЕКТ_ШАБЛОННАЯ_ФРАЗА}}", objectCondition)
    Call ReplacePlaceholderInWord(wdDoc, "{{ОБЪЕКТ_ШАБЛОННАЯ_ФРАЗА}}", objectPhrase)

    ' ---- 2. {{СОГЛАСОВАНИЕ_ИСТОЧНИКИ}} ----
    approvalTemplate = GetPhraseValue(wbMain, "{{СОГЛАСОВАНИЕ_ИСТОЧНИКИ}}", dealType)

    If UCase(dealType) = "ДДУ" Then
        approvalInnerPlaceholder = "{{ДАННЫЕ_ПРОТОКОЛА}}"
    ElseIf UCase(dealType) = "ДКП" Then
        approvalInnerPlaceholder = "{{ДАННЫЕ_ПИСЬМА}}"
    Else
        approvalInnerPlaceholder = ""
    End If

    If approvalInnerPlaceholder <> "" Then
        approvalData = GetDocRefValue(wbMain, approvalInnerPlaceholder, dealType, objectAddress)
        approvalPhrase = Replace(approvalTemplate, approvalInnerPlaceholder, approvalData)
    Else
        approvalPhrase = approvalTemplate
    End If

    Call ReplacePlaceholderInWord(wdDoc, "{{СОГЛАСОВАНИЕ_ИСТОЧНИКИ}}", approvalPhrase)

End Sub


Function GetPhraseValue(wbMain As Workbook, placeholderText As String, conditionText As String) As String

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    Set ws = wbMain.Worksheets("Phrases")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow
        If Trim(CStr(ws.Cells(r, 1).Value)) = Trim(placeholderText) _
           And Trim(CStr(ws.Cells(r, 2).Value)) = Trim(conditionText) Then

            GetPhraseValue = CStr(ws.Cells(r, 3).Value)
            Exit Function
        End If
    Next r

    GetPhraseValue = ""

End Function


Function NormalizeAddress(ByVal s As String) As String

    s = Trim(s)
    s = Replace(s, Chr(160), " ")
    s = Replace(s, "ё", "е")
    s = Replace(s, "Ё", "Е")
    s = Replace(s, "—", "-")
    s = Replace(s, "–", "-")

    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop

    NormalizeAddress = UCase(s)

End Function


Function GetDocRefValue(wbMain As Workbook, placeholderText As String, conditionText As String, objectAddress As String) As String

    Dim wsDocRefs As Worksheet
    Dim lastRow As Long
    Dim r As Long

    Dim srcSheet As String
    Dim addrColLetter As String
    Dim noColLetter As String
    Dim dateColLetter As String

    Dim filePath As String
    Dim wbExt As Workbook
    Dim wsExt As Worksheet

    Dim addrColIdx As Long
    Dim noColIdx As Long
    Dim dateColIdx As Long

    Dim lastDataRow As Long
    Dim i As Long

    Dim currentAddr As String
    Dim normTarget As String
    Dim dateVal As Variant
    Dim noVal As String

    Set wsDocRefs = wbMain.Worksheets("DocRefs")
    lastRow = wsDocRefs.Cells(wsDocRefs.Rows.Count, 1).End(xlUp).Row
    normTarget = NormalizeAddress(objectAddress)

    For r = 2 To lastRow
        If Trim(CStr(wsDocRefs.Cells(r, 1).Value)) = Trim(placeholderText) _
           And Trim(CStr(wsDocRefs.Cells(r, 2).Value)) = Trim(conditionText) Then

            srcSheet = Trim(CStr(wsDocRefs.Cells(r, 3).Value))
            addrColLetter = Trim(CStr(wsDocRefs.Cells(r, 4).Value))
            noColLetter = Trim(CStr(wsDocRefs.Cells(r, 5).Value))
            dateColLetter = Trim(CStr(wsDocRefs.Cells(r, 6).Value))
            Exit For
        End If
    Next r

    If srcSheet = "" Then
        GetDocRefValue = ""
        Exit Function
    End If

    If UCase(conditionText) = "ДДУ" Then
        filePath = GetSourcePathByNumber(wbMain, "0.4")
    ElseIf UCase(conditionText) = "ДКП" Then
        filePath = GetSourcePathByNumber(wbMain, "0.3")
    Else
        GetDocRefValue = ""
        Exit Function
    End If

    If filePath = "" Then
        GetDocRefValue = ""
        Exit Function
    End If

    Set wbExt = Workbooks.Open(Filename:=filePath, ReadOnly:=True, UpdateLinks:=False)
    Set wsExt = wbExt.Worksheets(srcSheet)

    addrColIdx = ColLetterToIndex(addrColLetter)
    noColIdx = ColLetterToIndex(noColLetter)
    dateColIdx = ColLetterToIndex(dateColLetter)

    lastDataRow = wsExt.Cells(wsExt.Rows.Count, addrColIdx).End(xlUp).Row

    For i = 2 To lastDataRow
        currentAddr = NormalizeAddress(CStr(wsExt.Cells(i, addrColIdx).Value))

        If currentAddr = normTarget Then

            If UCase(conditionText) = "ДДУ" Then
                GetDocRefValue = Trim(CStr(wsExt.Cells(i, noColIdx).Value))
            ElseIf UCase(conditionText) = "ДКП" Then
                dateVal = wsExt.Cells(i, dateColIdx).Value
                noVal = Trim(CStr(wsExt.Cells(i, noColIdx).Value))

                If IsDate(dateVal) Then
                    GetDocRefValue = Format(CDate(dateVal), "dd.mm.yyyy") & " № " & noVal
                Else
                    GetDocRefValue = CStr(dateVal) & " № " & noVal
                End If
            End If

            wbExt.Close SaveChanges:=False
            Exit Function
        End If
    Next i

    wbExt.Close SaveChanges:=False
    GetDocRefValue = ""

End Function


' ============================================================
' TABLE 1 — APPX_111a: округ / район / адрес (простые скаляры)
' ============================================================
Sub FillScalars_APPX_111a(wbMain As Workbook, wbData As Workbook, wdDoc As Object)

    Dim placeholders As Variant
    Dim i As Long
    Dim ph As String
    Dim v As String

    placeholders = Array("{{ОКРУГ}}", "{{РАЙОН}}", "{{АДРЕС}}")

    For i = LBound(placeholders) To UBound(placeholders)
        ph = CStr(placeholders(i))
        v = GetDataScalarByPlaceholder(wbMain, wbData, ph)
        Call ReplacePlaceholderInWord(wdDoc, ph, v)
    Next i

End Sub


' ============================================================
' Считает активные категории квартир (сумма > 0) по плейсхолдерам
' {{1К}}...{{СТ}}. Используется и в Table 2, и в подвале Table 3.
' ============================================================
Function GetActiveCategories(wbMain As Workbook, wbData As Workbook) As Object

    Dim result As Object
    Dim catCodes As Variant
    Dim i As Long
    Dim ph As String
    Dim v As String

    Set result = CreateObject("Scripting.Dictionary")

    catCodes = Array("1К", "2К", "3К", "4К", "5К", "1Е", "2Е", "3Е", "4Е", "5Е", "СТ")

    For i = LBound(catCodes) To UBound(catCodes)
        ph = "{{" & CStr(catCodes(i)) & "}}"
        v = GetDataScalarByPlaceholder(wbMain, wbData, ph)

        If Trim(v) <> "" Then
            If IsNumeric(v) Then
                If CDbl(v) <> 0 Then
                    result.Add CStr(catCodes(i)), CDbl(v)
                End If
            End If
        End If
    Next i

    Set GetActiveCategories = result

End Function


' ============================================================
' TABLE 2 — APPX_111b: список категорий квартир с суммами.
' Структура таблицы: 11 строк, 3 столбца, БЕЗ merge.
' Столбец 2 (Cell(r,2)) содержит плейсхолдер {{1К}} и т.д.
' Неактивные категории — строка удаляется целиком.
' ============================================================
Sub FillTable_APPX_111b(wdDoc As Object, activeCats As Object)

    Dim wdTable As Object
    Dim r As Long
    Dim catCode As String
    Dim ph As String
    Dim cellText As String

    Set wdTable = wdDoc.Tables(2)

    For r = wdTable.Rows.Count To 1 Step -1
        cellText = SafeGetCellText(wdTable, r, 2)
        catCode = ExtractCategoryCode(cellText)

        If catCode <> "" Then
            If Not activeCats.Exists(catCode) Then
                wdTable.Rows(r).Delete
            Else
                ph = "{{" & catCode & "}}"
                Call ReplacePlaceholderInWord(wdDoc, ph, Format(activeCats(catCode), "#,##0"))
            End If
        End If
    Next r

End Sub


' ============================================================
' TABLE 3 — APPX_112: аналоги-аукционы.
' Реальная структура:
'   Row 1-2  — двухуровневая шапка, 11 грид-колонок (не трогаем)
'   Row 3    — строка-образец для клонирования данных (11 колонок)
'   Rows 4.. — подвал "ИТОГО релевантная цена..." по категориям
' ============================================================
Sub FillTable_APPX_112(wbMain As Workbook, wbData As Workbook, wdDoc As Object, activeCats As Object)

    Dim wsMapper As Worksheet
    Dim wsData As Worksheet
    Dim wdTable As Object

    Dim lastMapRow As Long
    Dim r As Long
    Dim mapFound As Boolean

    Dim excelSheetName As String
    Dim keyColLetter As String
    Dim startRow As Long
    Dim keyColIdx As Long
    Dim lastDataRow As Long
    Dim rowCount As Long

    Dim i As Long
    Dim wordRow As Long

    Dim cat As Variant
    Dim ph As String
    Dim rowText As String
    Dim catCode As String

    Const SAMPLE_ROW As Long = 3

    Set wsMapper = wbMain.Worksheets("Mapper")
    lastMapRow = wsMapper.Cells(wsMapper.Rows.Count, 2).End(xlUp).Row
    mapFound = False

    ' Берём базовые параметры блока APPX_112 из первой строки блока
    For r = 2 To lastMapRow
        If Trim(CStr(wsMapper.Cells(r, 2).Value)) = "APPX_112" Then
            mapFound = True
            excelSheetName = Trim(CStr(wsMapper.Cells(r, 5).Value))
            keyColLetter = Trim(CStr(wsMapper.Cells(r, 6).Value))
            startRow = CLng(wsMapper.Cells(r, 7).Value)
            Exit For
        End If
    Next r

    If Not mapFound Then
        MsgBox "В Mapper нет строк для APPX_112."
        Exit Sub
    End If

    Set wsData = wbData.Worksheets(excelSheetName)
    keyColIdx = ColLetterToIndex(keyColLetter)
    lastDataRow = wsData.Cells(wsData.Rows.Count, keyColIdx).End(xlUp).Row

    If lastDataRow < startRow Then
        MsgBox "В листе " & excelSheetName & " нет данных для APPX_112."
        Exit Sub
    End If

    rowCount = lastDataRow - startRow + 1

    Set wdTable = wdDoc.Tables(3)

    ' =======================================================
    ' ШАГ 1. Чистим footer:
    ' удаляем только те строки, где есть плейсхолдер категории,
    ' но этой категории нет в activeCats
    ' =======================================================
    For r = wdTable.Rows.Count To 1 Step -1

        rowText = SafeGetRowText(wdTable, r)
        catCode = ExtractCategoryCode(rowText)

        If catCode <> "" Then
            If Not activeCats.Exists(catCode) Then
                wdTable.Rows(r).Delete
            End If
        End If
    Next r

    ' =======================================================
    ' ШАГ 2. Добавляем строки под контент:
    ' одна строка уже есть в шаблоне (строка 3),
    ' если данных больше — вставляем новые строки после неё
    ' =======================================================
    If rowCount > 1 Then
        Call InsertEmptyRowsAfter(wdTable, SAMPLE_ROW, rowCount - 1)
    End If

    ' =======================================================
    ' ШАГ 3. Заполняем контентную часть построчно,
    ' начиная со строки 3
    ' =======================================================
    For i = 0 To rowCount - 1
        wordRow = SAMPLE_ROW + i

        For r = 2 To lastMapRow
            If Trim(CStr(wsMapper.Cells(r, 2).Value)) = "APPX_112" Then

                Dim colOrder As Long
                Dim colLetter As String
                Dim sheetName As String
                Dim valueRow As Long
                Dim colIdx As Long

                colOrder = CLng(wsMapper.Cells(r, 3).Value)
                sheetName = Trim(CStr(wsMapper.Cells(r, 5).Value))
                colLetter = Trim(CStr(wsMapper.Cells(r, 6).Value))

                If colLetter <> "" Then
                    Set wsData = wbData.Worksheets(sheetName)
                    colIdx = ColLetterToIndex(colLetter)
                    valueRow = startRow + i

                    Call SafeSetCellText(wdTable, wordRow, colOrder, wsData.Cells(valueRow, colIdx).Text)
                End If
            End If
        Next r
    Next i

    ' =======================================================
    ' ШАГ 4. Подставляем суммы в оставшиеся footer-строки
    ' =======================================================
    For Each cat In activeCats.Keys
        ph = "{{" & CStr(cat) & "}}"
        Call ReplacePlaceholderInWord(wdDoc, ph, Format(activeCats(cat), "#,##0"))
    Next cat

    ' =======================================================
    ' ШАГ 5. Подставляем количество аналогов-строк в {{КВАРТИРЫ_СЧЕТ}}
    ' rowCount здесь — это именно число строк данных, залитых в Table 3
    ' =======================================================
    Call ReplacePlaceholderInWord(wdDoc, "{{КВАРТИРЫ_СЧЕТ}}", CStr(rowCount))

    
End Sub


' ============================================================
' TABLE 4 — APPX_113a: матрица корректировки (коэффициенты).
' ============================================================
Sub FillTable_APPX_113a(wbMain As Workbook, wbData As Workbook, wdDoc As Object)

    Dim wsMapper As Worksheet
    Dim wsData As Worksheet
    Dim wdTable As Object

    Dim lastMapRow As Long
    Dim r As Long
    Dim mapFound As Boolean

    Dim excelSheetName As String
    Dim startRow As Long
    Dim colLetter As String

    Dim rowCount As Long
    Dim i As Long
    Dim wordRow As Long
    Dim colOrder As Long

    Dim footerSum As String
    Dim cellRef As String
    Dim sheetName As String

    Const SAMPLE_ROW As Long = 2

    Set wsMapper = wbMain.Worksheets("Mapper")
    lastMapRow = wsMapper.Cells(wsMapper.Rows.Count, 2).End(xlUp).Row
    Set wdTable = wdDoc.Tables(4)

    rowCount = 0
    For r = 2 To lastMapRow
        If Trim(CStr(wsMapper.Cells(r, 2).Value)) = "APPX_113a" Then
            rowCount = rowCount + 1
        End If
    Next r

    If rowCount = 0 Then
        MsgBox "В Mapper нет строк для APPX_113a."
        Exit Sub
    End If

    For r = 2 To lastMapRow
        If Trim(CStr(wsMapper.Cells(r, 2).Value)) = "APPX_113a" Then
            mapFound = True
            excelSheetName = Trim(CStr(wsMapper.Cells(r, 5).Value))
            startRow = CLng(wsMapper.Cells(r, 7).Value)
            Exit For
        End If
    Next r

    If Not mapFound Then
        MsgBox "В Mapper нет базовой строки для APPX_113a."
        Exit Sub
    End If

    Set wsData = wbData.Worksheets(excelSheetName)

    If wdTable.Rows.Count < SAMPLE_ROW + rowCount - 1 Then
        MsgBox "В таблице APPX_113a не хватает строк под данные Mapper."
        Exit Sub
    End If

    ' Колонка Word определяется по порядку строк блока APPX_113a в Mapper
    colOrder = 0
    For r = 2 To lastMapRow
        If Trim(CStr(wsMapper.Cells(r, 2).Value)) = "APPX_113a" Then
            colOrder = colOrder + 1
            colLetter = Trim(CStr(wsMapper.Cells(r, 6).Value))

            Dim colIdx As Long
            Dim valueRow As Long

            colIdx = ColLetterToIndex(colLetter)

            For i = 0 To rowCount - 1
                wordRow = SAMPLE_ROW + i
                valueRow = startRow + i
                Call SafeSetCellText(wdTable, wordRow, colOrder, wsData.Cells(valueRow, colIdx).Text)
            Next i
        End If
    Next r

    footerSum = ""

    For r = 2 To lastMapRow
        If Trim(CStr(wsMapper.Cells(r, 8).Value)) = "{{113A_SUM}}" Then
            sheetName = Trim(CStr(wsMapper.Cells(r, 5).Value))
            cellRef = Trim(CStr(wsMapper.Cells(r, 6).Value))

            If sheetName <> "" And cellRef <> "" Then
                Set wsData = wbData.Worksheets(sheetName)
                footerSum = wsData.Range(cellRef).Text
                Exit For
            End If
        End If
    Next r

    Call ReplacePlaceholderInWord(wdDoc, "{{113A_SUM}}", footerSum)

End Sub


' ============================================================
' TABLE 5 — APPX_113b: ценообразование по квартирам + 3 суммы.
' Таблица содержит вертикальный merge в шапке (Rows 1-2),
' поэтому вставка строк идёт через Selection.InsertRowsBelow.
' ============================================================
Sub FillTable_APPX_113b(wbMain As Workbook, wbData As Workbook, wdDoc As Object)

    Dim wsMapper As Worksheet
    Dim wsData As Worksheet
    Dim wdTable As Object

    Dim lastMapRow As Long
    Dim r As Long
    Dim mapFound As Boolean

    Dim excelSheetName As String
    Dim keyColLetter As String
    Dim startRow As Long
    Dim keyColIdx As Long
    Dim lastDataRow As Long
    Dim rowCount As Long

    Dim i As Long
    Dim wordRow As Long

    Dim footerRow1 As Long
    Dim footerRow2 As Long
    Dim footerRow3 As Long

    Dim sumZ As Double
    Dim sumAF As Double
    Dim sumAI As Double

    Const SAMPLE_ROW As Long = 3
    Const FOOTER_VALUE_COL As Long = 3

    Set wsMapper = wbMain.Worksheets("Mapper")
    lastMapRow = wsMapper.Cells(wsMapper.Rows.Count, 2).End(xlUp).Row
    mapFound = False

    For r = 2 To lastMapRow
        If Trim(CStr(wsMapper.Cells(r, 2).Value)) = "APPX_113b" Then
            mapFound = True
            excelSheetName = Trim(CStr(wsMapper.Cells(r, 5).Value))
            keyColLetter = Trim(CStr(wsMapper.Cells(r, 6).Value))
            startRow = CLng(wsMapper.Cells(r, 7).Value)
            Exit For
        End If
    Next r

    If Not mapFound Then
        MsgBox "В Mapper нет строк для APPX_113b."
        Exit Sub
    End If

    Set wsData = wbData.Worksheets(excelSheetName)
    keyColIdx = ColLetterToIndex(keyColLetter)
    lastDataRow = wsData.Cells(wsData.Rows.Count, keyColIdx).End(xlUp).Row

    If lastDataRow < startRow Then Exit Sub

    rowCount = lastDataRow - startRow + 1

    Set wdTable = wdDoc.Tables(5)

    ' ---- Вставляем (rowCount - 1) новых строк ПОСЛЕ строки-образца ----
    If rowCount > 1 Then
        For i = 1 To rowCount - 1
            wdTable.Cell(SAMPLE_ROW, 1).Range.Select
            wdDoc.Application.Selection.InsertRowsBelow 1
        Next i
    End If

    ' ---- Заполняем построчно, начиная со строки-образца ----
    For i = 0 To rowCount - 1
        wordRow = SAMPLE_ROW + i

        For r = 2 To lastMapRow
            If Trim(CStr(wsMapper.Cells(r, 2).Value)) = "APPX_113b" Then

                Dim colOrder As Long
                Dim colLetter As String
                Dim sheetName As String
                Dim valueRow As Long
                Dim colIdx As Long

                colOrder = CLng(wsMapper.Cells(r, 3).Value)
                sheetName = Trim(CStr(wsMapper.Cells(r, 5).Value))
                colLetter = Trim(CStr(wsMapper.Cells(r, 6).Value))

                If colLetter <> "" Then
                    Set wsData = wbData.Worksheets(sheetName)
                    colIdx = ColLetterToIndex(colLetter)
                    valueRow = startRow + i

                    Call SafeSetCellText(wdTable, wordRow, colOrder, wsData.Cells(valueRow, colIdx).Text)
                End If
            End If
        Next r
    Next i

    ' ---- Подвал: 3 строки сразу после данных ----
    footerRow1 = SAMPLE_ROW + rowCount
    footerRow2 = footerRow1 + 1
    footerRow3 = footerRow2 + 1

    sumZ = SumColumnByLetter(wsData, "Z", startRow, lastDataRow)
    sumAF = SumColumnByLetter(wsData, "AF", startRow, lastDataRow)
    sumAI = SumColumnByLetter(wsData, "AI", startRow, lastDataRow)

    Call SafeSetCellText(wdTable, footerRow1, FOOTER_VALUE_COL, Format(sumZ, "#,##0"))
    Call SafeSetCellText(wdTable, footerRow2, FOOTER_VALUE_COL, Format(sumAF, "#,##0"))
    Call SafeSetCellText(wdTable, footerRow3, FOOTER_VALUE_COL, Format(sumAI, "#,##0"))

End Sub


' ============================================================
' TABLE 6 — APPX_114: варианты первоначального взноса.
' ============================================================
Sub FillTable_APPX_114(wbMain As Workbook, wbData As Workbook, wdDoc As Object)

    Dim wsMapper As Worksheet
    Dim wsData As Worksheet
    Dim wdTable As Object

    Dim lastMapRow As Long
    Dim r As Long
    Dim mapFound As Boolean

    Dim excelSheetName As String
    Dim keyColLetter As String
    Dim startRow As Long
    Dim keyColIdx As Long
    Dim lastDataRow As Long
    Dim rowCount As Long

    Dim i As Long
    Dim wordRow As Long

    Const SAMPLE_ROW As Long = 3

    Set wsMapper = wbMain.Worksheets("Mapper")
    lastMapRow = wsMapper.Cells(wsMapper.Rows.Count, 2).End(xlUp).Row
    mapFound = False

    For r = 2 To lastMapRow
        If Trim(CStr(wsMapper.Cells(r, 2).Value)) = "APPX_114" Then
            mapFound = True
            excelSheetName = Trim(CStr(wsMapper.Cells(r, 5).Value))
            keyColLetter = Trim(CStr(wsMapper.Cells(r, 6).Value))
            startRow = CLng(wsMapper.Cells(r, 7).Value)
            Exit For
        End If
    Next r

    If Not mapFound Then
        MsgBox "В Mapper нет строк для APPX_114."
        Exit Sub
    End If

    Set wsData = wbData.Worksheets(excelSheetName)
    keyColIdx = ColLetterToIndex(keyColLetter)
    lastDataRow = wsData.Cells(wsData.Rows.Count, keyColIdx).End(xlUp).Row

    If lastDataRow < startRow Then Exit Sub

    rowCount = lastDataRow - startRow + 1

    Set wdTable = wdDoc.Tables(6)

    If rowCount > 1 Then
        Call InsertEmptyRowsAfter(wdTable, SAMPLE_ROW, rowCount - 1)
    End If

    For i = 0 To rowCount - 1
        wordRow = SAMPLE_ROW + i

        For r = 2 To lastMapRow
            If Trim(CStr(wsMapper.Cells(r, 2).Value)) = "APPX_114" Then

                Dim colOrder As Long
                Dim colLetter As String
                Dim sheetName As String
                Dim valueRow As Long
                Dim colIdx As Long

                colOrder = CLng(wsMapper.Cells(r, 3).Value)
                sheetName = Trim(CStr(wsMapper.Cells(r, 5).Value))
                colLetter = Trim(CStr(wsMapper.Cells(r, 6).Value))

                Set wsData = wbData.Worksheets(sheetName)
                colIdx = ColLetterToIndex(colLetter)
                valueRow = startRow + i

                Call SafeSetCellText(wdTable, wordRow, colOrder, wsData.Cells(valueRow, colIdx).Text)
            End If
        Next r
    Next i

End Sub
