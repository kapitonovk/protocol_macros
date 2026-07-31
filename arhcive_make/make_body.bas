Attribute VB_Name = "make_body"
Option Explicit

' Делает ОДНО тело протокола для строки Sources с заданным Номером
' и возвращает путь к созданному docx-файлу.
Function MakeBodyFileForSource(wbMain As Workbook, wdApp As Object, sourceNumber As Long) As String

    Dim wsSources As Worksheet
    Dim wbData As Workbook
    Dim wdBody As Object

    Dim bodyPath As String
    Dim dataPath As String
    Dim lastRow As Long
    Dim i As Long
    Dim numVal As Variant

    Dim resultPath As String

    Set wsSources = wbMain.Worksheets("Sources")

    lastRow = wsSources.Cells(wsSources.Rows.Count, 1).End(xlUp).Row

    ' 1. Путь к шаблону BODY (Номер = 0.2)
    For i = 2 To lastRow
        numVal = wsSources.Cells(i, 1).Value
        If CStr(numVal) = "0.2" Then
            bodyPath = wsSources.Cells(i, 4).Value
            Exit For
        End If
    Next i

    ' 2. Путь к data-файлу для раздела 1 (Номер = 1)
    For i = 2 To lastRow
        numVal = wsSources.Cells(i, 1).Value
        If CStr(numVal) = "1" Then
            dataPath = wsSources.Cells(i, 4).Value
            Exit For
        End If
    Next i

    If bodyPath = "" Or dataPath = "" Then
        MsgBox "Не найден путь к template.body (0.2) или data1.xlsx (1) в Sources."
        MakeBodyFileForSource = ""
        Exit Function
    End If

    ' 3. Открываем data-файл
    Set wbData = Workbooks.Open(dataPath)

    ' 4. Создаём документ BODY по шаблону 0.2
    Set wdBody = wdApp.Documents.Add(Template:=bodyPath)

    ' 5. Таблица квартир BODY_MAIN
    Call FillTable_BODY_MAIN(wbMain, wbData, wdBody)

    ' 6. Скаляры: протокол, даты, список приложений, адрес
    Call ApplyScalarsToDocument_BODY(wbMain, wbData, wdBody, sourceNumber)

    ' 7. Сохраняем тело в файл и закрываем data
    resultPath = wbMain.Path & "\BODY_" & CStr(sourceNumber) & ".docx"
    wdBody.SaveAs2 Filename:=resultPath
    wbData.Close SaveChanges:=False
    wdBody.Close SaveChanges:=True

    MakeBodyFileForSource = resultPath

End Function


' Отдельный тестовый запуск: соберёт тело для Номер = 1
Sub RunBodyOnly()

    Dim wbMain As Workbook
    Dim wdApp As Object
    Dim bodyFile As String

    Set wbMain = ThisWorkbook

    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = True

    bodyFile = MakeBodyFileForSource(wbMain, wdApp, 1)

    Debug.Print "BODY file: " & bodyFile

End Sub


' ============================================================
' Таблица BODY_MAIN (матрица квартир) в template.body
' ============================================================
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
    lastMapRow = wsMapper.Cells(wsMapper.Rows.Count, 1).End(xlUp).Row
    firstRow = 0

    For r = 2 To lastMapRow
        If wsMapper.Cells(r, 2).Value = "BODY_MAIN" Then
            firstRow = r
            excelSheetName = wsMapper.Cells(r, 5).Value    ' ExcelSheet
            keyColLetter = wsMapper.Cells(r, 6).Value      ' ExcelColumnRef (буква)
            startRow = CLng(wsMapper.Cells(r, 7).Value)    ' StartRow
            Exit For
        End If
    Next r

    If firstRow = 0 Then
        MsgBox "В Mapper нет строк для BODY_MAIN."
        Exit Sub
    End If

    ' Определяем длину данных по ключевой колонке
    Set wsData = wbData.Worksheets(excelSheetName)
    keyColIdx = Body_ColLetterToIndex(keyColLetter)

    lastDataRow = wsData.Cells(wsData.Rows.Count, keyColIdx).End(xlUp).Row
    If lastDataRow < startRow Then
        MsgBox "В листе " & excelSheetName & " нет данных для BODY_MAIN."
        Exit Sub
    End If

    rowCount = lastDataRow - startRow + 1   ' сколько строк заливать

    ' В новом template.body таблица квартир — первая таблица (index = 1)[file:344]
    Set wdTable = wdBody.Tables(1)

    ' Чистим старые строки данных (оставляем шапку + строку-образец)
    Do While wdTable.Rows.Count > 2
        wdTable.Rows(3).Delete
    End If

    ' Убедимся, что строка-образец (2) существует
    If wdTable.Rows.Count < 2 Then
        wdTable.Rows.Add
    End If

    ' Добавляем строки под все данные:
    ' первая строка данных пойдёт в уже существующую строку 2,
    ' поэтому клонируем строку-образец (row 2) ещё (rowCount - 1) раз
    For i = 1 To rowCount - 1
        wdTable.Rows(2).Range.Copy
        wdTable.Rows.Add
    Next i

    ' Заполняем таблицу
    ' В Word: строка 1 – шапка, с 2 – данные
    For i = 0 To rowCount - 1
        wordRow = 2 + i

        ' Остальные колонки по Mapper’у (BODY_MAIN)
        For r = 2 To lastMapRow
            If wsMapper.Cells(r, 2).Value = "BODY_MAIN" Then

                Dim colOrder As Long
                Dim colLetter As String
                Dim sheetName As String
                Dim valueRow As Long
                Dim colIdx As Long

                colOrder = CLng(wsMapper.Cells(r, 3).Value)    ' WordColumnOrder
                sheetName = wsMapper.Cells(r, 5).Value          ' ExcelSheet
                colLetter = wsMapper.Cells(r, 6).Value          ' ExcelColumnRef

                Set wsData = wbData.Worksheets(sheetName)
                colIdx = Body_ColLetterToIndex(colLetter)
                valueRow = startRow + i

                wdTable.Cell(wordRow, colOrder).Range.Text = wsData.Cells(valueRow, colIdx).Text
            End If
        Next r
    Next i

End Sub


' ============================================================
' Скаляры для BODY: протокол, даты, список приложений, адрес
' ============================================================
Sub ApplyScalarsToDocument_BODY(wbMain As Workbook, wbData As Workbook, wdDoc As Object, sourceNumber As Long)

    Dim sourceRow As Long
    Dim protNum As String
    Dim date1 As String
    Dim date1Formatted As String
    Dim date2 As String
    Dim appendixList As String
    Dim addressVal As String

    sourceRow = Body_FindSourceRowByNumber(wbMain, sourceNumber)

    If sourceRow = 0 Then
        MsgBox "Не найдена строка в Sources для Номер = " & CStr(sourceNumber)
        Exit Sub
    End If

    protNum = Body_GetConfigValue(wbMain, "{{ПРОТ_НОМЕР}}")
    date1 = Body_GetConfigValue(wbMain, "{{ДАТА1}}")
    date1Formatted = Body_FormatDateForProtocol(date1)

    ' Пока логика простая:
    ' дата2 = дата1 + 90 дней, если дата1 — корректная дата
    If IsDate(date1) Then
        date2 = Format(DateAdd("d", 90, CDate(date1)), "dd.mm.yyyy")
    Else
        date2 = date1
    End If

    appendixList = Body_GetSourcesValue(wbMain, sourceRow, "Список приложений")
    addressVal = Body_GetDataScalarByPlaceholder_BODY(wbMain, wbData, "{{АДРЕС}}")

    Call Body_ReplacePlaceholderInWord(wdDoc, "{{ПРОТ_НОМЕР}}", protNum)
    Call Body_ReplacePlaceholderInWord(wdDoc, "{{ДАТА1}}", date1)
    Call Body_ReplacePlaceholderInWord(wdDoc, "{{ДАТА1ФОРМАТ}}", date1Formatted)
    Call Body_ReplacePlaceholderInWord(wdDoc, "{{ДАТА2}}", date2)
    Call Body_ReplacePlaceholderInWord(wdDoc, "{{ПРИЛОЖЕНИЯ}}", appendixList)
    Call Body_ReplacePlaceholderInWord(wdDoc, "{{АДРЕС}}", addressVal)

End Sub


' ============================================================
' УТИЛИТЫ make_body (именованы с префиксом Body_, чтобы не конфликтовать с make_appx)
' ============================================================

Function Body_ColLetterToIndex(colLetter As String) As Long
    colLetter = UCase(Trim(colLetter))
    Body_ColLetterToIndex = Range(colLetter & "1").Column
End Function


Function Body_FindHeaderColumn(ws As Worksheet, headerText As String) As Long
    Dim lastCol As Long, c As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        If Trim(CStr(ws.Cells(1, c).Value)) = Trim(headerText) Then
            Body_FindHeaderColumn = c
            Exit Function
        End If
    Next c
    Body_FindHeaderColumn = 0
End Function


Function Body_FindSourceRowByNumber(wbMain As Workbook, sourceNumber As Variant) As Long
    Dim ws As Worksheet, lastRow As Long, numCol As Long, r As Long
    Set ws = wbMain.Worksheets("Sources")
    numCol = Body_FindHeaderColumn(ws, "Номер")
    If numCol = 0 Then
        Body_FindSourceRowByNumber = 0
        Exit Function
    End If
    lastRow = ws.Cells(ws.Rows.Count, numCol).End(xlUp).Row
    For r = 2 To lastRow
        If Trim(CStr(ws.Cells(r, numCol).Value)) = Trim(CStr(sourceNumber)) Then
            Body_FindSourceRowByNumber = r
            Exit Function
        End If
    Next r
    Body_FindSourceRowByNumber = 0
End Function


Function Body_GetConfigValue(wbMain As Workbook, placeholderText As String) As String
    Dim ws As Worksheet, lastRow As Long, r As Long
    Set ws = wbMain.Worksheets("Config")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastRow
        If Trim(CStr(ws.Cells(r, 2).Value)) = Trim(placeholderText) Then
            Body_GetConfigValue = CStr(ws.Cells(r, 3).Value)
            Exit Function
        End If
    Next r
    Body_GetConfigValue = ""
End Function


Function Body_GetSourcesValue(wbMain As Workbook, sourceRow As Long, columnHeader As String) As String
    Dim ws As Worksheet, colNum As Long
    Set ws = wbMain.Worksheets("Sources")
    colNum = Body_FindHeaderColumn(ws, columnHeader)
    If colNum = 0 Then
        Body_GetSourcesValue = ""
        Exit Function
    End If
    Body_GetSourcesValue = CStr(ws.Cells(sourceRow, colNum).Value)
End Function


' Берёт скаляр по плейсхолдеру из Mapper0 (Placeholder, Sheet, Address, Start row)[file:343]
Function Body_GetDataScalarByPlaceholder_BODY(wbMain As Workbook, wbData As Workbook, placeholderText As String) As String

    Dim wsMap0 As Worksheet
    Dim lastRow As Long, r As Long
    Dim sheetName As String
    Dim addrRef As String
    Dim startRowVal As Variant
    Dim wsData As Worksheet
    Dim colIdx As Long

    Set wsMap0 = wbMain.Worksheets("Mapper0")
    lastRow = wsMap0.Cells(wsMap0.Rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow
        If Trim(CStr(wsMap0.Cells(r, 1).Value)) = Trim(placeholderText) Then

            sheetName = Trim(CStr(wsMap0.Cells(r, 2).Value))
            addrRef = Trim(CStr(wsMap0.Cells(r, 3).Value))
            startRowVal = wsMap0.Cells(r, 4).Value

            If sheetName <> "" Then
                Set wsData = wbData.Worksheets(sheetName)

                ' Если указан прямой адрес ячейки (например "B3")
                If addrRef <> "" And (InStr(addrRef, "1") > 0 Or InStr(addrRef, "2") > 0 Or InStr(addrRef, "3") > 0 Or InStr(addrRef, "4") > 0 Or InStr(addrRef, "5") > 0 Or InStr(addrRef, "6") > 0 Or InStr(addrRef, "7") > 0 Or InStr(addrRef, "8") > 0 Or InStr(addrRef, "9") > 0 Or InStr(addrRef, "0") > 0) Then
                    Body_GetDataScalarByPlaceholder_BODY = CStr(wsData.Range(addrRef).Value)
                    Exit Function
                End If

                ' Если задана только буква столбца и StartRow (как для {{АДРЕС_Т}})
                If addrRef <> "" And IsNumeric(startRowVal) Then
                    colIdx = Body_ColLetterToIndex(addrRef)
                    Body_GetDataScalarByPlaceholder_BODY = CStr(wsData.Cells(CLng(startRowVal), colIdx).Value)
                    Exit Function
                End If
            End If
        End If
    Next r

    Body_GetDataScalarByPlaceholder_BODY = ""

End Function


Sub Body_ReplacePlaceholderInWord(wdDoc As Object, findText As String, replaceText As String)
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


Function Body_FormatDateForProtocol(dateText As String) As String
    Dim d As Date, monthName As String
    If IsDate(dateText) Then
        d = CDate(dateText)
        monthName = LCase(Format(d, "mmmm"))
        Body_FormatDateForProtocol = "«" & Format(d, "dd") & "» " & monthName & " " & Format(d, "yyyy")
    Else
        Body_FormatDateForProtocol = dateText
    End If
End Function
