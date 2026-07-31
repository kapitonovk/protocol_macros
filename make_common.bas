Attribute VB_Name = "make_common"
Option Explicit

Function GetConfigValue(wbMain As Workbook, placeholderText As String) As String
    Dim ws As Worksheet, lastRow As Long, r As Long
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


Function FormatDateForProtocol(dateText As String) As String
    Dim d As Date, monthName As String
    If IsDate(dateText) Then
        d = CDate(dateText)
        monthName = LCase(Format(d, "mmmm"))
        FormatDateForProtocol = Chr(171) & Format(d, "dd") & Chr(187) & " " & monthName & " " & Format(d, "yyyy")
    Else
        FormatDateForProtocol = dateText
    End If
End Function


Function FindHeaderColumn(ws As Worksheet, headerText As String) As Long
    Dim lastCol As Long, c As Long
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
    Dim ws As Worksheet, lastRow As Long, numCol As Long, r As Long
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


Function GetSourcePathByNumber_Common(wbMain As Workbook, sourceNumber As String) As String
    Dim ws As Worksheet, sourceRow As Long, colPath As Long
    Set ws = wbMain.Worksheets("Sources")
    sourceRow = FindSourceRowByNumber(wbMain, sourceNumber)
    colPath = FindHeaderColumn(ws, "Путь")
    If sourceRow = 0 Or colPath = 0 Then
        GetSourcePathByNumber_Common = ""
    Else
        GetSourcePathByNumber_Common = Trim(CStr(ws.Cells(sourceRow, colPath).Value))
    End If
End Function


Function GetSourcesValue(wbMain As Workbook, sourceRow As Long, columnHeader As String) As String
    Dim ws As Worksheet, colNum As Long
    Set ws = wbMain.Worksheets("Sources")
    colNum = FindHeaderColumn(ws, columnHeader)
    If colNum = 0 Then
        GetSourcesValue = ""
        Exit Function
    End If
    GetSourcesValue = CStr(ws.Cells(sourceRow, colNum).Value)
End Function


Function GetDataScalarByPlaceholder(wbMain As Workbook, wbData As Workbook, placeholderText As String) As String
    Dim wsMapper As Worksheet, lastRow As Long, r As Long
    Dim sheetName As String, cellRef As String, wsData As Worksheet
    Set wsMapper = wbMain.Worksheets("Mapper")
    lastRow = wsMapper.Cells(wsMapper.Rows.Count, 2).End(xlUp).Row
    For r = 2 To lastRow
        If Trim(CStr(wsMapper.Cells(r, 8).Value)) = Trim(placeholderText) Then
            sheetName = Trim(CStr(wsMapper.Cells(r, 5).Value))
            cellRef = Trim(CStr(wsMapper.Cells(r, 6).Value))
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


Function ColLetterToIndex(colLetter As String) As Long
    colLetter = UCase(Trim(colLetter))
    ColLetterToIndex = Range(colLetter & "1").Column
End Function
