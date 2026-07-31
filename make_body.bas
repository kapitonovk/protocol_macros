Attribute VB_Name = "make_body"
Option Explicit

' ============================================================
' ИЗМЕНЕНИЯ ОТНОСИТЕЛЬНО ПРЕДЫДУЩЕЙ ВЕРСИИ:
'  1) шаблон body берётся по Sources 0.2 (было 0.1 — это титульник);
'  2) таблица квартир — wdDoc.Tables(1) (в шаблоне body таблица одна);
'  3) появилась параметризованная BuildBodyDoc(...) для вызова из make_report;
'  4) вставка строк через InsertEmptyRowsAfter + SafeSetCellText (make_appx),
'     заполняется колонка 1 "№ п.п.";
'  5) удалены дубли процедур, которые уже есть в make_appx
'     (ColLetterToIndex, FindHeaderColumn, FindSourceRowByNumber,
'      GetDataScalarByPlaceholder, ReplacePlaceholderInWord) —
'     иначе вызов из третьего модуля даёт "Ambiguous name";
'  6) GetConfigValue нормализует дату к dd.mm.yyyy.
' ============================================================


' ============================================================
' ВЫЗЫВАЕТСЯ ИЗ make_report: заполняет один экземпляр body
' ============================================================
Public Sub BuildBodyDoc(wbMain As Workbook, wbData As Workbook, wdDoc As Object, _
                        ByVal sourceNumber As Variant, ByVal sourceRow As Long)

    Call FillTable_BODY_MAIN(wbMain, wbData, wdDoc)
    Call ApplyScalarsToDocument(wbMain, wbData, wdDoc, sourceNumber)

End Sub


' ============================================================
' ОТЛАДОЧНАЯ ТОЧКА ВХОДА: только body по источнику №1
' ============================================================
Sub BuildSection_1_1_BODY_MAIN()

    Dim wbMain As Workbook
    Dim wbData As Workbook
    Dim wdApp As Object
    Dim wdBody As Object

    Dim bodyPath As String
    Dim dataPath As String
    Dim sourceRow As Long

    Set wbMain = ThisWorkbook

    bodyPath = GetSourcePathByNumber(wbMain, "0.2")   ' template.body.docx
    dataPath = GetSourcePathByNumber(wbMain, "1")     ' data1.xlsx
    sourceRow = FindSourceRowByNumber(wbMain, "1")

    If bodyPath = "" Or dataPath = "" Then
        MsgBox "Не найден путь к шаблону BODY (0.2) или data-файлу (1) в Sources."
        Exit Sub
    End If

    Set wbData = Workbooks.Open(Filename:=dataPath, ReadOnly:=True, UpdateLinks:=False)

    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = True
    Set wdBody = wdApp.Documents.Add(Template:=bodyPath)

    Call BuildBodyDoc(wbMain, wbData, wdBody, "1", sourceRow)

    wbData.Close SaveChanges:=False

End Sub


' ============================================================
' ТАБЛИЦА КВАРТИР BODY_MAIN (Tables(1) шаблона body)
' Строка 1 — шапка, строка 2 — образец данных.
' Колонка 1 — "№ п.п." (нумерация), колонки 2.. — по Mapper.
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

    Const SAMPLE_ROW As Long = 2

    Set wsMapper = wbMain.Worksheets("Mapper")

    lastMapRow = wsMapper.Cells(wsMapper.Rows.Count, 1).End(xlUp).Row
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

    ' В шаблоне body одна таблица: 1 — шапка, 2 — строка-образец
    Set wdTable = wdBody.Tables(1)

    ' Убираем лишние строки данных, если они были в шаблоне
    Do While wdTable.Rows.Count > SAMPLE_ROW
        wdTable.Rows(wdTable.Rows.Count).Delete
    Loop

    If wdTable.Rows.Count < SAMPLE_ROW Then
        wdTable.Rows.Add
    End If

    ' Добавляем строки под данные (первая строка данных = строка-образец)
    If rowCount > 1 Then
        Call InsertEmptyRowsAfter(wdTable, SAMPLE_ROW, rowCount - 1)
    End If

    ' Заполняем
    For i = 0 To rowCount - 1
        wordRow = SAMPLE_ROW + i

        ' Колонка 1: № п.п.
        Call SafeSetCellText(wdTable, wordRow, 1, CStr(i + 1))

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

                If colLetter <> "" Then
                    Set wsData = wbData.Worksheets(sheetName)
                    colIdx = ColLetterToIndex(colLetter)
                    valueRow = startRow + i

                    Call SafeSetCellText(wdTable, wordRow, colOrder, wsData.Cells(valueRow, colIdx).Text)
                End If
            End If
        Next r
    Next i

End Sub


' ============================================================
' СКАЛЯРЫ BODY: номер/даты/приложения/адрес
' ============================================================
Sub ApplyScalarsToDocument(wbMain As Workbook, wbData As Workbook, wdDoc As Object, sourceNumber As Variant)

    Dim sourceRow As Long
    Dim protNum As String
    Dim date1 As String
    Dim date1Formatted As String
    Dim date2 As String
    Dim appendixList As String
    Dim addressVal As String

    sourceRow = FindSourceRowByNumber(wbMain, sourceNumber)

    If sourceRow = 0 Then
        MsgBox "Не найдена строка в Sources для Номер = " & CStr(sourceNumber)
        Exit Sub
    End If

    protNum = GetConfigValue(wbMain, "{{ПРОТ_НОМЕР}}")
    date1 = GetConfigValue(wbMain, "{{ДАТА1}}")
    date1Formatted = FormatDateForProtocol(date1)

    If IsDate(date1) Then
        date2 = Format(DateAdd("d", 90, CDate(date1)), "dd.mm.yyyy")
    Else
        date2 = date1
    End If

    appendixList = GetSourcesValue(wbMain, sourceRow, "Список приложений")
    addressVal = GetDataScalarByPlaceholder(wbMain, wbData, "{{АДРЕС}}")

    Call ReplacePlaceholderInWord(wdDoc, "{{ПРОТ_НОМЕР}}", protNum)
    Call ReplacePlaceholderInWord(wdDoc, "{{ДАТА1ФОРМАТ}}", date1Formatted)
    Call ReplacePlaceholderInWord(wdDoc, "{{ДАТА1}}", date1)
    Call ReplacePlaceholderInWord(wdDoc, "{{ДАТА2}}", date2)
    Call ReplacePlaceholderInWord(wdDoc, "{{ПРИЛОЖЕНИЯ}}", appendixList)
    Call ReplacePlaceholderInWord(wdDoc, "{{АДРЕС}}", addressVal)

End Sub


' ============================================================
' CONFIG: значение по плейсхолдеру (даты приводим к dd.mm.yyyy)
' ============================================================
Function GetConfigValue(wbMain As Workbook, placeholderText As String) As String

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim v As Variant

    Set ws = wbMain.Worksheets("Config")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For r = 2 To lastRow
        If Trim(CStr(ws.Cells(r, 2).Value)) = Trim(placeholderText) Then

            v = ws.Cells(r, 3).Value

            If IsDate(v) Then
                GetConfigValue = Format(CDate(v), "dd.mm.yyyy")
            Else
                GetConfigValue = Trim(CStr(v))
            End If

            Exit Function
        End If
    Next r

    GetConfigValue = ""

End Function


' ============================================================
' SOURCES: значение по заголовку колонки в заданной строке
' ============================================================
Function GetSourcesValue(wbMain As Workbook, sourceRow As Long, columnHeader As String) As String

    Dim ws As Worksheet
    Dim colNum As Long

    Set ws = wbMain.Worksheets("Sources")
    colNum = FindHeaderColumn(ws, columnHeader)

    If colNum = 0 Then
        GetSourcesValue = ""
        Exit Function
    End If

    GetSourcesValue = Trim(CStr(ws.Cells(sourceRow, colNum).Value))

End Function


' ============================================================
' «01» января 2000
' ============================================================
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
