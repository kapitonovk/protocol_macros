Option Explicit

' ============================================================
' √ÀŒ¡¿À‹Õ¿ﬂ —“–” “”–¿ œ”“≈…   —À”∆≈¡Õ€Ã ‘¿…À¿Ã (0.1 - 0.6)
' ============================================================
Private Type ServicePaths
    TemplateTitle As String        ' 0.1 - template.title.docx
    TemplateBody As String         ' 0.2 - template.body.docx
    TemplateApproval As String     ' 0.3 - template.approval.docx
    TemplateAppendix As String     ' 0.4 - template.appendix.docx
    PoolLettersFile As String      ' 0.5
    DduDataFile As String          ' 0.6
End Type

Private Type ReportObject
    ObjectNumber As Long
    DataPath As String
    Address As String
    TotalSum As Double
End Type


' ============================================================
' √À¿¬Õ¿ﬂ “Œ◊ ¿ ¬’Œƒ¿
' ============================================================
Sub RunFullReport()

    Dim wbMain As Workbook
    Dim paths As ServicePaths
    Dim objects() As ReportObject
    Dim objCount As Long

    Dim wbPoolLetters As Workbook
    Dim wbDduData As Workbook

    Dim wdApp As Object

    Dim tmpFolder As String
    Dim titlePath As String
    Dim approvalPath As String
    Dim bodyPaths() As String
    Dim appxPaths() As String

    Dim i As Long
    Dim finalPath As String

    Set wbMain = ThisWorkbook

    Call ParseSourcesConfig(wbMain, paths)
    objects = GetObjectsList(wbMain)
    objCount = 0

    On Error Resume Next
    objCount = UBound(objects) + 1
    On Error GoTo 0

    If objCount = 0 Then
        MsgBox "¬ Sources ÌÂ Ì‡È‰ÂÌÓ ÌË Ó‰ÌÓ„Ó Ó·˙ÂÍÚ‡ (ÒÚÓÍË Ò ˆÂÎ˚Ï ˜ËÒÎÓÏ ‚ ÍÓÎÓÌÍÂ ÕÓÏÂ)."
        Exit Sub
    End If

    If Not ValidateServicePaths(paths) Then
        Exit Sub
    End If

    tmpFolder = wbMain.Path & "\_tmp_report_build\"
    If Dir(tmpFolder, vbDirectory) = "" Then
        MkDir tmpFolder
    End If

    Set wbPoolLetters = Workbooks.Open(Filename:=paths.PoolLettersFile, ReadOnly:=True, UpdateLinks:=False)
    Set wbDduData = Workbooks.Open(Filename:=paths.DduDataFile, ReadOnly:=True, UpdateLinks:=False)

    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = True

    ReDim bodyPaths(1 To objCount)
    ReDim appxPaths(1 To objCount)

    ' ---- “ÂÎÓ + ÔËÎÓÊÂÌËÂ ‰Îˇ Í‡Ê‰Ó„Ó Ó·˙ÂÍÚ‡ (‡‰ÂÒ/ÒÛÏÏ‡ ÌÛÊÌ˚ Á‡‡ÌÂÂ ‰Îˇ approval) ----
    For i = 1 To objCount

        Dim wbData As Workbook
        Set wbData = Workbooks.Open(Filename:=objects(i - 1).DataPath, ReadOnly:=True, UpdateLinks:=False)

        bodyPaths(i) = tmpFolder & "_tmp_body_" & i & ".docx"
        Call BuildOneObjectBody(wbMain, wbData, wdApp, paths.TemplateBody, objects(i - 1).ObjectNumber, bodyPaths(i))

        appxPaths(i) = tmpFolder & "_tmp_appx_" & i & ".docx"
        Call BuildOneObjectAppendix(wbMain, wbData, wdApp, wbPoolLetters, wbDduData, _
                                     paths.TemplateAppendix, objects(i - 1).ObjectNumber, appxPaths(i), _
                                     objects(i - 1).Address, objects(i - 1).TotalSum)

        wbData.Close SaveChanges:=False

    Next i

    ' ---- “ËÚÛÎ¸ÌËÍ (ÔÓÒÎÂ ˆËÍÎ‡, Ú.Í. ÒÂÈ˜‡Ò ËÒÔÓÎ¸ÁÛÂÚ ÚÓÎ¸ÍÓ Config) ----
    titlePath = tmpFolder & "_tmp_title.docx"
    Call BuildTitlePage(wbMain, wdApp, paths.TemplateTitle, objects, objCount, titlePath)

    ' ---- ÀËÒÚ ÛÚ‚ÂÊ‰ÂÌËˇ ----
    approvalPath = tmpFolder & "_tmp_approval.docx"
    Call BuildApprovalPage(wbMain, wdApp, paths.TemplateApproval, objects, objCount, approvalPath)

    ' ---- ‘ËÌ‡Î¸Ì‡ˇ ÒÍÎÂÈÍ‡: title -> body*N -> approval -> appx*N ----
    finalPath = wbMain.Path & "\report_result_" & Format(Now, "yyyymmdd_hhnnss") & ".docx"
    Call AssembleFinalDocument(wdApp, titlePath, bodyPaths, appxPaths, approvalPath, objCount, finalPath)

    wbPoolLetters.Close SaveChanges:=False
    wbDduData.Close SaveChanges:=False

    MsgBox "ŒÚ˜∏Ú ÒÓ·‡Ì: " & finalPath

End Sub


' ============================================================
' –‡Á·Ó Sources: ÒÎÛÊÂ·Ì˚Â ÔÛÚË (0.1 - 0.6)
' ============================================================
Sub ParseSourcesConfig(wbMain As Workbook, ByRef paths As ServicePaths)

    Dim ws As Worksheet
    Dim numCol As Long
    Dim pathCol As Long
    Dim lastRow As Long
    Dim r As Long
    Dim numText As String

    Set ws = wbMain.Worksheets("Sources")

    numCol = FindHeaderColumn(ws, "ÕÓÏÂ")
    pathCol = FindHeaderColumn(ws, "œÛÚ¸")

    If numCol = 0 Or pathCol = 0 Then
        MsgBox "Õ‡ ÎËÒÚÂ Sources ÌÂ Ì‡È‰ÂÌ˚ Á‡„ÓÎÓ‚ÍË 'ÕÓÏÂ' Ë/ËÎË 'œÛÚ¸'."
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
    If paths.PoolLettersFile = "" Then missing = missing & "0.5 (ÔËÒ¸Ï‡ œ”ÀÓ‚) " & vbCrLf
    If paths.DduDataFile = "" Then missing = missing & "0.6 (ddu_data) " & vbCrLf

    If missing <> "" Then
        MsgBox "¬ Sources ÌÂ Ì‡È‰ÂÌ˚ ÔÛÚË ‰Îˇ:" & vbCrLf & missing
        ValidateServicePaths = False
    Else
        ValidateServicePaths = True
    End If

End Function


' ============================================================
' –‡Á·Ó Sources: ÒÔËÒÓÍ Ó·˙ÂÍÚÓ‚ (ˆÂÎ˚Â ÌÓÏÂ‡ 1, 2, 3...)
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

    numCol = FindHeaderColumn(ws, "ÕÓÏÂ")
    pathCol = FindHeaderColumn(ws, "œÛÚ¸")

    lastRow = ws.Cells(ws.Rows.Count, numCol).End(xlUp).Row

    ReDim result(0 To lastRow)
    cnt = 0

    For r = 2 To lastRow
        numText = Trim(CStr(ws.Cells(r, numCol).Value))

        If IsNumeric(numText) Then
            If InStr(numText, ".") = 0 Then
                result(cnt).ObjectNumber = CLng(numText)
                result(cnt).DataPath = Trim(CStr(ws.Cells(r, pathCol).Value))
                result(cnt).Address = ""
                result(cnt).TotalSum = 0
                cnt = cnt + 1
            End If
        End If
    Next r

    If cnt = 0 Then
        ReDim result(-1 To -1)
        GetObjectsList = result
        Exit Function
    End If

    ReDim Preserve result(0 To cnt - 1)

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
' Œ¡Ÿ»≈ ”“»À»“€
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
        FormatDateForProtocol = "´" & Format(d, "dd") & "ª " & monthName & " " & Format(d, "yyyy")
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
    s = Replace(s, "∏", "Â")
    s = Replace(s, "®", "≈")
    s = Replace(s, "ó", "-")
    s = Replace(s, "ñ", "-")

    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop

    NormalizeAddress = UCase(s)

End Function


Function FindSourceRowByNumber(wbMain As Workbook, sourceNumber As Long) As Long

    Dim ws As Worksheet
    Dim numCol As Long
    Dim lastRow As Long
    Dim r As Long

    Set ws = wbMain.Worksheets("Sources")
    numCol = FindHeaderColumn(ws, "ÕÓÏÂ")

    If numCol = 0 Then
        FindSourceRowByNumber = 0
        Exit Function
    End If

    lastRow = ws.Cells(ws.Rows.Count, numCol).End(xlUp).Row

    For r = 2 To lastRow
        If IsNumeric(ws.Cells(r, numCol).Value) Then
            If CLng(ws.Cells(r, numCol).Value) = sourceNumber Then
                FindSourceRowByNumber = r
                Exit Function
            End If
        End If
    Next r

    FindSourceRowByNumber = 0

End Function


Function GetSourcesValue(wbMain As Workbook, sourceRow As Long, columnHeader As String) As String

    Dim ws As Worksheet
    Dim colNum As Long

    Set ws = wbMain.Worksheets("Sources")
    colNum = FindHeaderColumn(ws, columnHeader)

    If colNum = 0 Or sourceRow = 0 Then
        GetSourcesValue = ""
        Exit Function
    End If

    GetSourcesValue = CStr(ws.Cells(sourceRow, colNum).Value)

End Function


' ============================================================
' “»“”À‹Õ»  (ËÒÔÓÎ¸ÁÛÂÚ Config: {{œ–Œ“_ÕŒÃ≈–}}, {{ƒ¿“¿1‘Œ–Ã¿“}})
' ============================================================
Sub BuildTitlePage(wbMain As Workbook, wdApp As Object, templatePath As String, _
                    objects() As ReportObject, objCount As Long, outPath As String)

    Dim wdDoc As Object
    Dim protNum As String
    Dim date1 As String
    Dim date1Formatted As String

    Set wdDoc = wdApp.Documents.Add(Template:=templatePath)

    protNum = GetConfigValue(wbMain, "{{œ–Œ“_ÕŒÃ≈–}}")
    date1 = GetConfigValue(wbMain, "{{ƒ¿“¿1}}")
    date1Formatted = FormatDateForProtocol(date1)

    Call ReplacePlaceholderInWord(wdDoc, "{{œ–Œ“_ÕŒÃ≈–}}", protNum)
    Call ReplacePlaceholderInWord(wdDoc, "{{ƒ¿“¿1‘Œ–Ã¿“}}", date1Formatted)

    wdDoc.SaveAs2 Filename:=outPath
    wdDoc.Close SaveChanges:=False

End Sub


' ============================================================
' “≈ÀŒ ŒƒÕŒ√Œ Œ¡⁄≈ “¿
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
        MsgBox "¬ Mapper ÌÂÚ ÒÚÓÍ ‰Îˇ BODY_MAIN."
        Exit Sub
    End If

    Set wsData = wbData.Worksheets(excelSheetName)
    keyColIdx = ColLetterToIndex(keyColLetter)

    lastDataRow = wsData.Cells(wsData.Rows.Count, keyColIdx).End(xlUp).Row
    If lastDataRow < startRow Then
        MsgBox "¬ ÎËÒÚÂ " & excelSheetName & " ÌÂÚ ‰‡ÌÌ˚ı ‰Îˇ BODY_MAIN."
        Exit Sub
    End If

    rowCount = lastDataRow - startRow + 1

    ' ¬ ÚÂÍÛ˘ÂÏ ¯‡·ÎÓÌÂ template.body Ú‡·ÎËˆ‡ Í‚‡ÚË ó œ≈–¬¿ﬂ (Ë Â‰ËÌÒÚ‚ÂÌÌ‡ˇ) Ú‡·ÎËˆ‡.
    Set wdTable = wdBody.Tables(1)

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

    Dim sourceRow As Long
    Dim appendixList As String

    protNum = GetConfigValue(wbMain, "{{œ–Œ“_ÕŒÃ≈–}}")
    date1 = GetConfigValue(wbMain, "{{ƒ¿“¿1}}")
    date1Formatted = FormatDateForProtocol(date1)

    If IsDate(date1) Then
        date2 = Format(DateAdd("d", 90, CDate(date1)), "dd.mm.yyyy")
    Else
        date2 = date1
    End If

    addressVal = GetDataScalarByPlaceholder(wbMain, wbData, "{{¿ƒ–≈—}}")

    sourceRow = FindSourceRowByNumber(wbMain, objectNumber)
    If sourceRow <> 0 Then
        appendixList = GetSourcesValue(wbMain, sourceRow, "—ÔËÒÓÍ ÔËÎÓÊÂÌËÈ")
    Else
        appendixList = ""
    End If

    Call ReplacePlaceholderInWord(wdDoc, "{{œ–Œ“_ÕŒÃ≈–}}", protNum)
    Call ReplacePlaceholderInWord(wdDoc, "{{ƒ¿“¿1}}", date1)
    Call ReplacePlaceholderInWord(wdDoc, "{{ƒ¿“¿1‘Œ–Ã¿“}}", date1Formatted)
    Call ReplacePlaceholderInWord(wdDoc, "{{ƒ¿“¿2}}", date2)
    Call ReplacePlaceholderInWord(wdDoc, "{{¿ƒ–≈—}}", addressVal)

    If appendixList <> "" Then
        Call ReplacePlaceholderInWord(wdDoc, "{{œ–»ÀŒ∆≈Õ»ﬂ}}", appendixList)
    End If

    Call ReplacePlaceholderInWord(wdDoc, "{{ÕŒÃ≈–_Œ¡⁄≈ “¿}}", CStr(objectNumber))

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
' œ–»ÀŒ∆≈Õ»≈ ŒƒÕŒ√Œ Œ¡⁄≈ “¿
' ============================================================
Sub BuildOneObjectAppendix(wbMain As Workbook, wbData As Workbook, wdApp As Object, _
                            wbPoolLetters As Workbook, wbDduData As Workbook, _
                            templatePath As String, objectNumber As Long, outPath As String, _
                            ByRef outAddress As String, ByRef outTotalSum As Double)

    Dim wdDoc As Object
    Dim activeCats As Object

    Set wdDoc = wdApp.Documents.Add(Template:=templatePath)

    Call FillScalars_APPX_111a(wbMain, wbData, wdDoc)

    outAddress = GetDataScalarByPlaceholder(wbMain, wbData, "{{¿ƒ–≈—}}")

    Call FillConditionalPhrases_APPX(wbMain, wbData, wdDoc, wbPoolLetters, wbDduData, outAddress)

    Set activeCats = GetActiveCategories(wbMain, wbData)

    Call FillTable_APPX_111b(wdDoc, activeCats)
    Call FillTable_APPX_112(wbMain, wbData, wdDoc, activeCats)
    Call FillTable_APPX_113a(wbMain, wbData, wdDoc)
    Call FillTable_APPX_113b(wbMain, wbData, wdDoc)
    Call FillTable_APPX_114(wbMain, wbData, wdDoc)

    Call ReplacePlaceholderInWord(wdDoc, "{{ÕŒÃ≈–_Œ¡⁄≈ “¿}}", CStr(objectNumber))

    Dim cat As Variant
    outTotalSum = 0
    For Each cat In activeCats.Keys
        outTotalSum = outTotalSum + activeCats(cat)
    Next cat

    wdDoc.SaveAs2 Filename:=outPath
    wdDoc.Close SaveChanges:=False

End Sub


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

    dealType = Trim(GetConfigValue(wbMain, "{{“»œ_—ƒ≈À »}}"))
    developerType = Trim(GetConfigValue(wbMain, "{{“»œ_«¿—“–Œ…Ÿ» ¿}}"))

    If UCase(dealType) = "ƒ œ" And LCase(developerType) = "ÙÓÌ‰" Then
        objectCondition = "ƒ œ_ÙÓÌ‰"
    Else
        objectCondition = "»ÌÓÂ"
    End If

    objectPhrase = GetPhraseValue(wbMain, "{{Œ¡⁄≈ “_ÿ¿¡ÀŒÕÕ¿ﬂ_‘–¿«¿}}", objectCondition)
    Call ReplacePlaceholderInWord(wdDoc, "{{Œ¡⁄≈ “_ÿ¿¡ÀŒÕÕ¿ﬂ_‘–¿«¿}}", objectPhrase)

    approvalTemplate = GetPhraseValue(wbMain, "{{—Œ√À¿—Œ¬¿Õ»≈_»—“Œ◊Õ» »}}", dealType)

    If UCase(dealType) = "ƒƒ”" Then
        approvalInnerPlaceholder = "{{ƒ¿ÕÕ€≈_œ–Œ“Œ ŒÀ¿}}"
    ElseIf UCase(dealType) = "ƒ œ" Then
        approvalInnerPlaceholder = "{{ƒ¿ÕÕ€≈_œ»—‹Ã¿}}"
    Else
        approvalInnerPlaceholder = ""
    End If

    If approvalInnerPlaceholder <> "" Then
        approvalData = LookupDocRefValue(wbMain, wbPoolLetters, wbDduData, dealType, objectAddress)
        approvalPhrase = Replace(approvalTemplate, approvalInnerPlaceholder, approvalData)
    Else
        approvalPhrase = approvalTemplate
    End If

    Call ReplacePlaceholderInWord(wdDoc, "{{—Œ√À¿—Œ¬¿Õ»≈_»—“Œ◊Õ» »}}", approvalPhrase)

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

    If UCase(conditionText) = "ƒƒ”" Then
        Set wsExt = wbDduData.Worksheets(1)
    ElseIf UCase(conditionText) = "ƒ œ" Then
        Set wsExt = wbPoolLetters.Worksheets(1)
    Else
        LookupDocRefValue = ""
        Exit Function
    End If

    addrColIdx = 1
    noColIdx = 2
    dateColIdx = 3

    lastDataRow = wsExt.Cells(wsExt.Rows.Count, addrColIdx).End(xlUp).Row

    For i = 2 To lastDataRow
        currentAddr = NormalizeAddress(CStr(wsExt.Cells(i, addrColIdx).Value))

        If currentAddr = normTarget Then

            dateVal = wsExt.Cells(i, dateColIdx).Value
            noVal = Trim(CStr(wsExt.Cells(i, noColIdx).Value))

            If IsDate(dateVal) Then
                LookupDocRefValue = Format(CDate(dateVal), "dd.mm.yyyy") & " π " & noVal
            Else
                LookupDocRefValue = CStr(dateVal) & " π " & noVal
            End If

            Exit Function
        End If
    Next i

    LookupDocRefValue = ""

End Function


Sub FillScalars_APPX_111a(wbMain As Workbook, wbData As Workbook, wdDoc As Object)

    Dim placeholders As Variant
    Dim i As Long
    Dim ph As String
    Dim v As String

    placeholders = Array("{{Œ –”√}}", "{{–¿…ŒÕ}}", "{{¿ƒ–≈—}}")

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

    catCodes = Array("1 ", "2 ", "3 ", "4 ", "5 ", "1≈", "2≈", "3≈", "4≈", "5≈", "—“")

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
        MsgBox "¬ Mapper ÌÂÚ ÒÚÓÍ ‰Îˇ APPX_112."
        Exit Sub
    End If

    Set wsData = wbData.Worksheets(excelSheetName)
    keyColIdx = ColLetterToIndex(keyColLetter)
    lastDataRow = wsData.Cells(wsData.Rows.Count, keyColIdx).End(xlUp).Row

    If lastDataRow < startRow Then
        MsgBox "¬ ÎËÒÚÂ " & excelSheetName & " ÌÂÚ ‰‡ÌÌ˚ı ‰Îˇ APPX_112."
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
        MsgBox "¬ Mapper ÌÂÚ ÒÚÓÍ ‰Îˇ APPX_113a."
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
        MsgBox "¬ Mapper ÌÂÚ ·‡ÁÓ‚ÓÈ ÒÚÓÍË ‰Îˇ APPX_113a."
        Exit Sub
    End If

    Set wsData = wbData.Worksheets(excelSheetName)

    If wdTable.Rows.Count < SAMPLE_ROW + rowCount - 1 Then
        MsgBox "¬ Ú‡·ÎËˆÂ APPX_113a ÌÂ ı‚‡Ú‡ÂÚ ÒÚÓÍ ÔÓ‰ ‰‡ÌÌ˚Â Mapper."
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
        MsgBox "¬ Mapper ÌÂÚ ÒÚÓÍ ‰Îˇ APPX_113b."
        Exit Sub
    End If

    Set wsData = wbData.Worksheets(excelSheetName)
    keyColIdx = ColLetterToIndex(keyColLetter)
    lastDataRow = wsData.Cells(wsData.Rows.Count, keyColIdx).End(xlUp).Row

    If lastDataRow < startRow Then Exit Sub

    rowCount = lastDataRow - startRow + 1

    Set wdTable = wdDoc.Tables(5)

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
        MsgBox "¬ Mapper ÌÂÚ ÒÚÓÍ ‰Îˇ APPX_114."
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
' À»—“ ”“¬≈–∆ƒ≈Õ»ﬂ (ËÒÔÓÎ¸ÁÛÂÚ Config: {{œ–Œ“_ÕŒÃ≈–}}, {{ƒ¿“¿1}})
' ============================================================
Sub BuildApprovalPage(wbMain As Workbook, wdApp As Object, templatePath As String, _
                       objects() As ReportObject, objCount As Long, outPath As String)

    Dim wdDoc As Object
    Dim protNum As String
    Dim date1 As String

    Set wdDoc = wdApp.Documents.Add(Template:=templatePath)

    protNum = GetConfigValue(wbMain, "{{œ–Œ“_ÕŒÃ≈–}}")
    date1 = GetConfigValue(wbMain, "{{ƒ¿“¿1}}")

    Call ReplacePlaceholderInWord(wdDoc, "{{œ–Œ“_ÕŒÃ≈–}}", protNum)
    Call ReplacePlaceholderInWord(wdDoc, "{{ƒ¿“¿1}}", date1)

    wdDoc.SaveAs2 Filename:=outPath
    wdDoc.Close SaveChanges:=False

End Sub


' ============================================================
' ‘»Õ¿À‹Õ¿ﬂ — À≈… ¿: title -> body(1..N) -> approval -> appendix(1..N)
' ============================================================
Sub AssembleFinalDocument(wdApp As Object, titlePath As String, bodyPaths() As String, _
                           appxPaths() As String, approvalPath As String, objCount As Long, finalPath As String)

    Dim wdFinal As Object
    Dim i As Long

    Set wdFinal = wdApp.Documents.Open(Filename:=titlePath)

    For i = 1 To objCount
        Call InsertFileWithPageBreak(wdFinal, bodyPaths(i))
    Next i

    Call InsertFileWithPageBreak(wdFinal, approvalPath)

    For i = 1 To objCount
        Call InsertFileWithPageBreak(wdFinal, appxPaths(i))
    Next i

    wdFinal.SaveAs2 Filename:=finalPath

End Sub


Sub InsertFileWithPageBreak(wdDoc As Object, otherPath As String)

    Dim endRange As Object

    Set endRange = wdDoc.Range(wdDoc.Content.End - 1, wdDoc.Content.End - 1)
    endRange.Collapse 0

    endRange.InsertBreak 7

    Set endRange = wdDoc.Range(wdDoc.Content.End - 1, wdDoc.Content.End - 1)
    endRange.Collapse 0
    endRange.InsertFile FileName:=otherPath

End Sub


Sub KillFolderContents(folderPath As String)

    Dim f As String
    f = Dir(folderPath & "*.docx")

    Do While f <> ""
        Kill folderPath & f
        f = Dir()
    Loop

End Sub
