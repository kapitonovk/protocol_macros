Option Explicit

' ============================================================
' ГЛОБАЛЬНАЯ СТРУКТУРА ПУТЕЙ К СЛУЖЕБНЫМ ФАЙЛАМ (0.1 - 0.6)
' Заполняется один раз в начале сборки отчёта
' ============================================================
Private Type ServicePaths
    TemplateTitle As String        ' 0.1 - template.title.docx
    TemplateBody As String         ' 0.2 - template.body.docx
    TemplateApproval As String     ' 0.3 - template.approval.docx
    TemplateAppendix As String     ' 0.4 - template.appendix.docx
    PoolLettersFile As String      ' 0.5 - ПОДБОР писем для ПУЛов.xlsx
    DduDataFile As String          ' 0.6 - ddu_data.xlsx
End Type

' Один объект отчёта = один data-файл (data_1.xlsx, data_2.xlsx, ...)
Private Type ReportObject
    ObjectNumber As Long           ' 1, 2, 3, ... — порядковый номер объекта
    DataPath As String             ' путь к data_N.xlsx
    Address As String              ' адрес объекта (заполняется в ходе сборки, для approval)
    TotalSum As Double             ' итоговая сумма объекта (заполняется в ходе сборки, для approval)
End Type


' ============================================================
' ГЛАВНАЯ ТОЧКА ВХОДА
' Собирает весь отчёт: титульник -> N тел -> лист утверждения -> N приложений
' ============================================================
Sub RunFullReport()

    Dim wbMain As Workbook
    Dim paths As ServicePaths
    Dim objects() As ReportObject
    Dim objCount As Long

    Dim wbPoolLetters As Workbook   ' общий справочник, открывается один раз
    Dim wbDduData As Workbook       ' общий справочник, открывается один раз

    Dim wdApp As Object

    Dim tmpFolder As String
    Dim titlePath As String
    Dim approvalPath As String
    Dim bodyPaths() As String
    Dim appxPaths() As String

    Dim i As Long
    Dim finalPath As String

    Set wbMain = ThisWorkbook

    ' ---- Шаг 0: разбираем Sources на служебные пути и список объектов ----
    Call ParseSourcesConfig(wbMain, paths)
    objects = GetObjectsList(wbMain)
    objCount = 0

    ' Определяем реальное количество объектов (массив может прийти с запасом)
    On Error Resume Next
    objCount = UBound(objects) + 1
    On Error GoTo 0

    If objCount = 0 Then
        MsgBox "В Sources не найдено ни одного объекта (строки с целым числом в колонке Номер)."
        Exit Sub
    End If

    ' Проверяем, что все служебные пути найдены
    If Not ValidateServicePaths(paths) Then
        Exit Sub
    End If

    ' ---- Папка для временных файлов сборки ----
    tmpFolder = wbMain.Path & "\_tmp_report_build\"
    If Dir(tmpFolder, vbDirectory) = "" Then
        MkDir tmpFolder
    End If

    ' ---- Открываем общие справочники ОДИН раз на весь цикл ----
    ' ПРЕДПОЛОЖЕНИЕ: оба файла открываются ReadOnly, структура читается функциями
    ' GetDocRefValue-подобными ниже. Если реальная структура файлов отличается —
    ' поправить нужно только внутри LookupPoolLetter / LookupDduData.
    Set wbPoolLetters = Workbooks.Open(Filename:=paths.PoolLettersFile, ReadOnly:=True, UpdateLinks:=False)
    Set wbDduData = Workbooks.Open(Filename:=paths.DduDataFile, ReadOnly:=True, UpdateLinks:=False)

    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = True

    ReDim bodyPaths(1 To objCount)
    ReDim appxPaths(1 To objCount)

    ' ---- Шаг 2: титульник (один раз) ----
    titlePath = tmpFolder & "_tmp_title.docx"
    Call BuildTitlePage(wbMain, wdApp, paths.TemplateTitle, objects, objCount, titlePath)

    ' ---- Шаг 3: цикл по объектам — тело + приложение для каждого ----
    For i = 1 To objCount

        Dim wbData As Workbook
        Set wbData = Workbooks.Open(Filename:=objects(i - 1).DataPath, ReadOnly:=True, UpdateLinks:=False)

        ' -- Тело объекта --
        bodyPaths(i) = tmpFolder & "_tmp_body_" & i & ".docx"
        Call BuildOneObjectBody(wbMain, wbData, wdApp, paths.TemplateBody, objects(i - 1).ObjectNumber, bodyPaths(i))

        ' -- Приложение объекта --
        appxPaths(i) = tmpFolder & "_tmp_appx_" & i & ".docx"
        Call BuildOneObjectAppendix(wbMain, wbData, wdApp, wbPoolLetters, wbDduData, _
                                     paths.TemplateAppendix, objects(i - 1).ObjectNumber, appxPaths(i), _
                                     objects(i - 1).Address, objects(i - 1).TotalSum)

        ' Адрес и сумма объекта записываются функцией BuildOneObjectAppendix через ByRef,
        ' они понадобятся ниже для листа утверждения

        wbData.Close SaveChanges:=False

    Next i

    ' ---- Шаг 4: лист утверждения (один раз, после всех объектов) ----
    approvalPath = tmpFolder & "_tmp_approval.docx"
    Call BuildApprovalPage(wbMain, wdApp, paths.TemplateApproval, objects, objCount, approvalPath)

    ' ---- Шаг 5: финальная склейка в порядке title -> body*N -> approval -> appx*N ----
    finalPath = wbMain.Path & "\report_result_" & Format(Now, "yyyymmdd_hhnnss") & ".docx"
    Call AssembleFinalDocument(wdApp, titlePath, bodyPaths, appxPaths, approvalPath, objCount, finalPath)

    ' ---- Закрываем общие справочники ----
    wbPoolLetters.Close SaveChanges:=False
    wbDduData.Close SaveChanges:=False

    MsgBox "Отчёт собран: " & finalPath

    ' Временные файлы (_tmp_*.docx) оставлены в tmpFolder для отладки.
    ' Когда пайплайн будет обкатан — можно раскомментировать очистку:
    ' Call KillFolderContents(tmpFolder)

End Sub


' ============================================================
' ШАГ 0.1 — Разбор Sources: служебные пути (0.1 - 0.6)
' ============================================================
Sub ParseSourcesConfig(wbMain As Workbook, ByRef paths As ServicePaths)

    Dim ws As Worksheet
    Dim numCol As Long
    Dim pathCol As Long
    Dim lastRow As Long
    Dim r As Long
    Dim numText As String

    Set ws = wbMain.Worksheets("Sources")

    numCol = FindHeaderColumn(ws, "Номер")
    pathCol = FindHeaderColumn(ws, "Путь")

    If numCol = 0 Or pathCol = 0 Then
        MsgBox "На листе Sources не найдены заголовки 'Номер' и/или 'Путь'."
        Exit Sub
    End If

    lastRow = ws.Cells(ws.Rows.Count, numCol).End(xlUp).Row

    For r = 2 To lastRow
        numText = Trim(CStr(ws.Cells(r, numCol).Value))

        Select Case numText
            Case "0.1"
                paths.TemplateTitle = Trim(CStr(ws.Cells(r, pathCol).Value))
            Case "0.2"
                paths.TemplateBody = Trim(CStr(ws.Cells(r, pathCol).Value))
            Case "0.3"
                paths.TemplateApproval = Trim(CStr(ws.Cells(r, pathCol).Value))
            Case "0.4"
                paths.TemplateAppendix = Trim(CStr(ws.Cells(r, pathCol).Value))
            Case "0.5"
                paths.PoolLettersFile = Trim(CStr(ws.Cells(r, pathCol).Value))
            Case "0.6"
                paths.DduDataFile = Trim(CStr(ws.Cells(r, pathCol).Value))
        End Select
    Next r

End Sub


Function ValidateServicePaths(paths As ServicePaths) As Boolean

    Dim missing As String

    If paths.TemplateTitle = "" Then missing = missing & "0.1 (template.title) " & vbCrLf
    If paths.TemplateBody = "" Then missing = missing & "0.2 (template.body) " & vbCrLf
    If paths.TemplateApproval = "" Then missing = missing & "0.3 (template.approval) " & vbCrLf
    If paths.TemplateAppendix = "" Then missing = missing & "0.4 (template.appendix) " & vbCrLf
    If paths.PoolLettersFile = "" Then missing = missing & "0.5 (ПОДБОР писем для ПУЛов) " & vbCrLf
    If paths.DduDataFile = "" Then missing = missing & "0.6 (ddu_data) " & vbCrLf

    If missing <> "" Then
        MsgBox "В Sources не найдены пути для:" & vbCrLf & missing
        ValidateServicePaths = False
    Else
        ValidateServicePaths = True
    End If

End Function


' ============================================================
' ШАГ 0.2 — Разбор Sources: список объектов (целые номера 1, 2, 3...)
' Возвращает массив ReportObject, отсортированный по возрастанию номера
' ============================================================
Function GetObjectsList(wbMain As Workbook) As ReportObject()

    Dim ws As Worksheet
    Dim numCol As Long
    Dim pathCol As Long
    Dim lastRow As Long
    Dim r As Long
    Dim numText As String

    Dim result() As ReportObject
    Dim cnt As Long

    Set ws = wbMain.Worksheets("Sources")

    numCol = FindHeaderColumn(ws, "Номер")
    pathCol = FindHeaderColumn(ws, "Путь")

    lastRow = ws.Cells(ws.Rows.Count, numCol).End(xlUp).Row

    ReDim result(0 To lastRow) ' с запасом, обрежем по факту
    cnt = 0

    For r = 2 To lastRow
        numText = Trim(CStr(ws.Cells(r, numCol).Value))

        ' Объект — это строка, где Номер представляет собой ЦЕЛОЕ число (без точки).
        ' Служебные строки (0.1 и т.д.) содержат точку и сюда не попадают.
        If IsNumeric(numText) Then
            If InStr(numText, ".") = 0 Then
                result(cnt).ObjectNumber = CLng(numText)
                result(cnt).DataPath = Trim(CStr(ws.Cells(r, pathCol).Value))
                result(cnt).Address = ""     ' заполнится позже, во время сборки appendix
                result(cnt).TotalSum = 0      ' заполнится позже, во время сборки appendix
                cnt = cnt + 1
            End If
        End If
    Next r

    If cnt = 0 Then
        ReDim result(-1 To -1) ' пустой массив -> UBound + 1 = 0
        GetObjectsList = result
        Exit Function
    End If

    ReDim Preserve result(0 To cnt - 1)

    ' Сортировка по ObjectNumber (простая пузырьковая — объектов немного, скорости достаточно)
    Dim i As Long, j As Long
    Dim tmp As ReportObject

    For i = 0 To cnt - 2
        For j = 0 To cnt - 2 - i
            If result(j).ObjectNumber > result(j + 1).ObjectNumber Then
                tmp = result(j)
                result(j) = result(j + 1)
                result(j + 1) = tmp
            End If
        Next j
    Next i

    GetObjectsList = result

End Function


' ============================================================
' ОБЩИЕ УТИЛИТЫ (без изменений в логике — уже проверены ранее)
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


' ============================================================
' ШАГ 2 — ТИТУЛЬНИК (один раз на весь отчёт)
' ============================================================
Sub BuildTitlePage(wbMain As Workbook, wdApp As Object, templatePath As String, _
                    objects() As ReportObject, objCount As Long, outPath As String)

    Dim wdDoc As Object
    Dim addressList As String
    Dim i As Long

    Set wdDoc = wdApp.Documents.Add(Template:=templatePath)

    ' Собираем список адресов объектов через перенос строки —
    ' ПРЕДПОЛОЖЕНИЕ: в титульнике есть плейсхолдер {{СПИСОК_ОБЪЕКТОВ}}.
    ' Адреса на этом этапе ещё не известны (они читаются из data-файлов
    ' в BuildOneObjectAppendix), поэтому если титульник должен содержать
    ' реальные адреса — эту функцию нужно вызывать ПОСЛЕ цикла по объектам,
    ' а не до него. Сейчас для простоты титульник использует только
    ' общие реквизиты, не зависящие от объектов.

    Call ReplacePlaceholderInWord(wdDoc, "{{ДАТА_ОТЧЕТА}}", Format(Date, "dd.mm.yyyy"))
    Call ReplacePlaceholderInWord(wdDoc, "{{КОЛИЧЕСТВО_ОБЪЕКТОВ}}", CStr(objCount))

    wdDoc.SaveAs2 Filename:=outPath
    wdDoc.Close SaveChanges:=False

End Sub


' ============================================================
' ШАГ 3.1 — ТЕЛО ОДНОГО ОБЪЕКТА
' ============================================================
Sub BuildOneObjectBody(wbMain As Workbook, wbData As Workbook, wdApp As Object, _
                        templatePath As String, objectNumber As Long, outPath As String)

    Dim wdDoc As Object

    Set wdDoc = wdApp.Documents.Add(Template:=templatePath)

    Call FillTable_BODY_MAIN(wbMain, wbData, wdDoc)
    Call ApplyScalarsToDocument(wbMain, wbData, wdDoc, objectNumber)

    wdDoc.SaveAs2 Filename:=outPath
    wdDoc.Close SaveChanges:=False

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

    lastMapRow = wsMapper.Cells(wsMapper.Rows.Count, 2).End(xlUp).Row
    firstRow = 0

    For r = 2 To lastMapRow
        If Trim(CStr(wsMapper.Cells(r, 2).Value)) = "BODY_MAIN" Then
            firstRow = r
            excelSheetName = Trim(CStr(wsMapper.Cells(r, 5).Value))
            keyColLetter = Trim(CStr(wsMapper.Cells(r, 6).Value))
            startRow = CLng(wsMapper.Cells(r, 7).Value)
            Exit For
        End If
    Next r

    If firstRow = 0 Then
        MsgBox "В Mapper нет строк для BODY_MAIN."
        Exit Sub
    End If

    Set wsData = wbData.Worksheets(excelSheetName)
    keyColIdx = ColLetterToIndex(keyColLetter)

    lastDataRow = wsData.Cells(wsData.Rows.Count, keyColIdx).End(xlUp).Row
    If lastDataRow < startRow Then
        MsgBox "В листе " & excelSheetName & " нет данных для BODY_MAIN."
        Exit Sub
    End If

    rowCount = lastDataRow - startRow + 1

    Set wdTable = wdBody.Tables(2)

    Do While wdTable.Rows.Count > 2
        wdTable.Rows(3).Delete
    Loop

    If wdTable.Rows.Count < 2 Then
        wdTable.Rows.Add
    End If

    If rowCount > 1 Then
        Call InsertEmptyRowsAfter(wdTable, 2, rowCount - 1)
    End If

    For i = 0 To rowCount - 1
        wordRow = 2 + i

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


Sub ApplyScalarsToDocument(wbMain As Workbook, wbData As Workbook, wdDoc As Object, objectNumber As Long)

    Dim protNum As String
    Dim date1 As String
    Dim date1Formatted As String
    Dim date2 As String
    Dim addressVal As String

    protNum = GetConfigValue(wbMain, "{{ПРОТ_НОМЕР}}")
    date1 = GetConfigValue(wbMain, "{{ДАТА1}}")
    date1Formatted = FormatDateForProtocol(date1)

    If IsDate(date1) Then
        date2 = Format(DateAdd("d", 90, CDate(date1)), "dd.mm.yyyy")
    Else
        date2 = date1
    End If

    addressVal = GetDataScalarByPlaceholder(wbMain, wbData, "{{АДРЕС}}")

    Call ReplacePlaceholderInWord(wdDoc, "{{ПРОТ_НОМЕР}}", protNum)
    Call ReplacePlaceholderInWord(wdDoc, "{{ДАТА1}}", date1)
    Call ReplacePlaceholderInWord(wdDoc, "{{ДАТА1ФОРМАТ}}", date1Formatted)
    Call ReplacePlaceholderInWord(wdDoc, "{{ДАТА2}}", date2)
    Call ReplacePlaceholderInWord(wdDoc, "{{АДРЕС}}", addressVal)

    ' ПРЕДПОЛОЖЕНИЕ: в шаблоне body есть плейсхолдер {{НОМЕР_ОБЪЕКТА}},
    ' который отражает порядковый номер объекта в отчёте (1, 2, 3...).
    ' Если в реальном шаблоне плейсхолдер называется иначе — поменять здесь.
    Call ReplacePlaceholderInWord(wdDoc, "{{НОМЕР_ОБЪЕКТА}}", CStr(objectNumber))

End Sub


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


' ============================================================
' ШАГ 3.2 — ПРИЛОЖЕНИЕ ОДНОГО ОБЪЕКТА
' Возвращает адрес и итоговую сумму объекта через ByRef —
' они понадобятся ниже для листа утверждения (approval)
' ============================================================
Sub BuildOneObjectAppendix(wbMain As Workbook, wbData As Workbook, wdApp As Object, _
                            wbPoolLetters As Workbook, wbDduData As Workbook, _
                            templatePath As String, objectNumber As Long, outPath As String, _
                            ByRef outAddress As String, ByRef outTotalSum As Double)

    Dim wdDoc As Object
    Dim activeCats As Object

    Set wdDoc = wdApp.Documents.Add(Template:=templatePath)

    Call FillScalars_APPX_111a(wbMain, wbData, wdDoc)

    outAddress = GetDataScalarByPlaceholder(wbMain, wbData, "{{АДРЕС}}")

    Call FillConditionalPhrases_APPX(wbMain, wbData, wdDoc, wbPoolLetters, wbDduData, outAddress)

    Set activeCats = GetActiveCategories(wbMain, wbData)

    Call FillTable_APPX_111b(wdDoc, activeCats)
    Call FillTable_APPX_112(wbMain, wbData, wdDoc, activeCats)
    Call FillTable_APPX_113a(wbMain, wbData, wdDoc)
    Call FillTable_APPX_113b(wbMain, wbData, wdDoc)
    Call FillTable_APPX_114(wbMain, wbData, wdDoc)

    ' ПРЕДПОЛОЖЕНИЕ: в шаблоне appendix есть плейсхолдер {{НОМЕР_ОБЪЕКТА}}
    ' (например для подписи "Приложение к разделу N")
    Call ReplacePlaceholderInWord(wdDoc, "{{НОМЕР_ОБЪЕКТА}}", CStr(objectNumber))

    ' Итоговую сумму объекта берём из суммы активных категорий (Table 2/3) —
    ' это тот же набор данных, что уже считает GetActiveCategories.
    ' ПРЕДПОЛОЖЕНИЕ: "итоговая сумма объекта" = сумма всех активных категорий.
    ' Если реальная итоговая сумма берётся из другого места (например, из
    ' футера Table 5) — заменить эту строку на чтение нужной ячейки.
    Dim cat As Variant
    outTotalSum = 0
    For Each cat In activeCats.Keys
        outTotalSum = outTotalSum + activeCats(cat)
    Next cat

    wdDoc.SaveAs2 Filename:=outPath
    wdDoc.Close SaveChanges:=False

End Sub


' ============================================================
' УСЛОВНЫЕ ПЛЕЙСХОЛДЕРЫ ДЛЯ ПРИЛОЖЕНИЯ
' Теперь принимает уже ОТКРЫТЫЕ книги wbPoolLetters/wbDduData —
' они не открываются и не закрываются здесь, это делает RunFullReport
' ============================================================
Sub FillConditionalPhrases_APPX(wbMain As Workbook, wbData As Workbook, wdDoc As Object, _
                                 wbPoolLetters As Workbook, wbDduData As Workbook, objectAddress As String)

    Dim dealType As String
    Dim developerType As String

    Dim objectCondition As String
    Dim objectPhrase As String

    Dim approvalTemplate As String
    Dim approvalInnerPlaceholder As String
    Dim approvalData As String
    Dim approvalPhrase As String

    ' ПРЕДПОЛОЖЕНИЕ: "Тип" и "Застройщик" сейчас читаются из Config,
    ' а не из Sources — потому что в новой структуре Sources хранит только
    ' пути к файлам, а не атрибуты объекта. Если у объекта они должны
    ' различаться (не общие на весь отчёт) — эти два значения нужно
    ' переносить в отдельный лист объектов и читать оттуда.
    dealType = Trim(GetConfigValue(wbMain, "{{ТИП_СДЕЛКИ}}"))
    developerType = Trim(GetConfigValue(wbMain, "{{ТИП_ЗАСТРОЙЩИКА}}"))

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
        approvalData = LookupDocRefValue(wbMain, wbPoolLetters, wbDduData, dealType, objectAddress)
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


' ============================================================
' Ищет реквизиты документа (протокол / письмо) во ВНЕШНЕМ,
' уже открытом справочнике (wbDduData для ДДУ, wbPoolLetters для ДКП)
' по нормализованному адресу объекта.
'
' ПРЕДПОЛОЖЕНИЕ по структуре обоих файлов (проверить и поправить
' колонки при необходимости):
'   Колонка A — адрес объекта
'   Колонка B — номер документа
'   Колонка C — дата документа
' Лист — первый лист книги (ws(1)), если у тебя данные на другом листе —
' поменять "wb.Worksheets(1)" на конкретное имя листа.
' ============================================================
Function LookupDocRefValue(wbMain As Workbook, wbPoolLetters As Workbook, wbDduData As Workbook, _
                            conditionText As String, objectAddress As String) As String

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

    normTarget = NormalizeAddress(objectAddress)

    If UCase(conditionText) = "ДДУ" Then
        Set wsExt = wbDduData.Worksheets(1)      ' [ПРОВЕРИТЬ] правильный лист ddu_data.xlsx
    ElseIf UCase(conditionText) = "ДКП" Then
        Set wsExt = wbPoolLetters.Worksheets(1)  ' [ПРОВЕРИТЬ] правильный лист файла писем для ПУЛов
    Else
        LookupDocRefValue = ""
        Exit Function
    End If

    addrColIdx = 1   ' колонка A — адрес     [ПРОВЕРИТЬ]
    noColIdx = 2     ' колонка B — номер     [ПРОВЕРИТЬ]
    dateColIdx = 3   ' колонка C — дата      [ПРОВЕРИТЬ]

    lastDataRow = wsExt.Cells(wsExt.Rows.Count, addrColIdx).End(xlUp).Row

    For i = 2 To lastDataRow
        currentAddr = NormalizeAddress(CStr(wsExt.Cells(i, addrColIdx).Value))

        If currentAddr = normTarget Then

            dateVal = wsExt.Cells(i, dateColIdx).Value
            noVal = Trim(CStr(wsExt.Cells(i, noColIdx).Value))

            If IsDate(dateVal) Then
                LookupDocRefValue = Format(CDate(dateVal), "dd.mm.yyyy") & " № " & noVal
            Else
                LookupDocRefValue = CStr(dateVal) & " № " & noVal
            End If

            Exit Function
        End If
    Next i

    LookupDocRefValue = ""

End Function


' ============================================================
' TABLE 1 — APPX_111a
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

    For r = wdTable.Rows.Count To 1 Step -1
        rowText = SafeGetRowText(wdTable, r)
        catCode = ExtractCategoryCode(rowText)

        If catCode <> "" Then
            If Not activeCats.Exists(catCode) Then
                wdTable.Rows(r).Delete
            End If
        End If
    Next r

    If rowCount > 1 Then
        Call InsertEmptyRowsAfter(wdTable, SAMPLE_ROW, rowCount - 1)
    End If

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

    For Each cat In activeCats.Keys
        ph = "{{" & CStr(cat) & "}}"
        Call ReplacePlaceholderInWord(wdDoc, ph, Format(activeCats(cat), "#,##0"))
    Next cat

End Sub


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

    ' Вставка через Selection.InsertRowsBelow — сохранена намеренно из-за
    ' вертикального merge в шапке (Rows 1-2), см. риск в предыдущих версиях
    If rowCount > 1 Then
        Dim i2 As Long
        For i2 = 1 To rowCount - 1
            wdTable.Cell(SAMPLE_ROW, 1).Range.Select
            wdDoc.Application.Selection.InsertRowsBelow 1
        Next i2
    End If

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


' ============================================================
' ШАГ 4 — ЛИСТ УТВЕРЖДЕНИЯ (один раз, ПОСЛЕ всех тел, ПЕРЕД всеми приложениями)
' ПРЕДПОЛОЖЕНИЕ: в шаблоне approval есть Table(1) с шапкой (Row 1)
' и строкой-образцом (Row 2), содержащей плейсхолдеры {{АДРЕС}} и {{ИТОГО}}.
' Если структура другая — поправить SAMPLE_ROW и номера колонок.
' ============================================================
Sub BuildApprovalPage(wbMain As Workbook, wdApp As Object, templatePath As String, _
                       objects() As ReportObject, objCount As Long, outPath As String)

    Dim wdDoc As Object
    Dim wdTable As Object

    Dim i As Long
    Dim wordRow As Long

    Const SAMPLE_ROW As Long = 2
    Const COL_ADDRESS As Long = 1   ' [ПРОВЕРИТЬ] номер столбца с адресом
    Const COL_TOTAL As Long = 2     ' [ПРОВЕРИТЬ] номер столбца с итоговой суммой

    Dim grandTotal As Double

    Set wdDoc = wdApp.Documents.Add(Template:=templatePath)
    Set wdTable = wdDoc.Tables(1)

    If objCount > 1 Then
        Call InsertEmptyRowsAfter(wdTable, SAMPLE_ROW, objCount - 1)
    End If

    grandTotal = 0

    For i = 0 To objCount - 1
        wordRow = SAMPLE_ROW + i

        Call SafeSetCellText(wdTable, wordRow, COL_ADDRESS, objects(i).Address)
        Call SafeSetCellText(wdTable, wordRow, COL_TOTAL, Format(objects(i).TotalSum, "#,##0"))

        grandTotal = grandTotal + objects(i).TotalSum
    Next i

    ' ПРЕДПОЛОЖЕНИЕ: в шаблоне approval есть отдельный плейсхолдер
    ' {{ИТОГО_ПО_ВСЕМ_ОБЪЕКТАМ}} для общей суммы под таблицей
    Call ReplacePlaceholderInWord(wdDoc, "{{ИТОГО_ПО_ВСЕМ_ОБЪЕКТАМ}}", Format(grandTotal, "#,##0"))
    Call ReplacePlaceholderInWord(wdDoc, "{{КОЛИЧЕСТВО_ОБЪЕКТОВ}}", CStr(objCount))
    Call ReplacePlaceholderInWord(wdDoc, "{{ДАТА_ОТЧЕТА}}", Format(Date, "dd.mm.yyyy"))

    wdDoc.SaveAs2 Filename:=outPath
    wdDoc.Close SaveChanges:=False

End Sub


' ============================================================
' ШАГ 5 — ФИНАЛЬНАЯ СКЛЕЙКА
' Порядок: title -> body(1..N) -> approval -> appendix(1..N)
' ============================================================
Sub AssembleFinalDocument(wdApp As Object, titlePath As String, bodyPaths() As String, _
                           appxPaths() As String, approvalPath As String, objCount As Long, finalPath As String)

    Dim wdFinal As Object
    Dim i As Long

    ' Титульник становится основой финального документа
    Set wdFinal = wdApp.Documents.Open(Filename:=titlePath)

    ' -- Вставляем все тела --
    For i = 1 To objCount
        Call InsertFileWithPageBreak(wdFinal, bodyPaths(i))
    Next i

    ' -- Вставляем лист утверждения (между body и appendix) --
    Call InsertFileWithPageBreak(wdFinal, approvalPath)

    ' -- Вставляем все приложения --
    For i = 1 To objCount
        Call InsertFileWithPageBreak(wdFinal, appxPaths(i))
    Next i

    wdFinal.SaveAs2 Filename:=finalPath
    ' Финальный документ оставляем открытым, чтобы пользователь сразу увидел результат.
    ' Если нужно закрывать автоматически — раскомментировать:
    ' wdFinal.Close SaveChanges:=False

End Sub


' Вставляет содержимое файла otherPath в конец документа wdDoc,
' предварительно вставив разрыв страницы (кроме случая, когда
' документ полностью пуст — тогда разрыв не нужен)
Sub InsertFileWithPageBreak(wdDoc As Object, otherPath As String)

    Dim endRange As Object

    Set endRange = wdDoc.Range(wdDoc.Content.End - 1, wdDoc.Content.End - 1)
    endRange.Collapse 0 ' wdCollapseEnd

    ' Разрыв страницы перед вставляемым куском
    endRange.InsertBreak 7 ' wdPageBreak

    ' Пересчитываем конец документа после разрыва и вставляем файл
    Set endRange = wdDoc.Range(wdDoc.Content.End - 1, wdDoc.Content.End - 1)
    endRange.Collapse 0
    endRange.InsertFile FileName:=otherPath

End Sub


' ============================================================
' Вспомогательная очистка временных файлов (по желанию, сейчас не вызывается)
' ============================================================
Sub KillFolderContents(folderPath As String)

    Dim f As String
    f = Dir(folderPath & "*.docx")

    Do While f <> ""
        Kill folderPath & f
        f = Dir()
    Loop

End Sub
