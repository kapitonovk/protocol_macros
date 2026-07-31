Attribute VB_Name = "make_title"
Option Explicit

' Собирает титульный лист (шаблон 0.1) и возвращает готовый Word-документ,
' либо сохраняет его в файл (если передан savePath).
Function BuildTitlePage(wbMain As Workbook, wdApp As Object, Optional savePath As String = "") As Object

    Dim titlePath As String
    Dim protNum As String
    Dim date1 As String
    Dim date1Formatted As String
    Dim wdTitle As Object

    titlePath = GetSourcePathByNumber_Common(wbMain, "0.1")

    If titlePath = "" Then
        MsgBox "Не найден путь к шаблону title (0.1) в Sources."
        Exit Function
    End If

    Set wdTitle = wdApp.Documents.Add(Template:=titlePath)

    protNum = GetConfigValue(wbMain, "{{ПРОТ_НОМЕР}}")
    date1 = GetConfigValue(wbMain, "{{ДАТА1}}")
    date1Formatted = FormatDateForProtocol(date1)

    Call ReplacePlaceholderInWord(wdTitle, "{{ПРОТ_НОМЕР}}", protNum)
    Call ReplacePlaceholderInWord(wdTitle, "{{ДАТА1ФОРМАТ}}", date1Formatted)

    If savePath <> "" Then
        wdTitle.SaveAs2 Filename:=savePath
    End If

    Set BuildTitlePage = wdTitle

End Function


' Отдельный запуск для теста
Sub RunTitleOnly()

    Dim wbMain As Workbook
    Dim wdApp As Object
    Dim wdTitle As Object

    Set wbMain = ThisWorkbook

    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = True

    Set wdTitle = BuildTitlePage(wbMain, wdApp)

End Sub
