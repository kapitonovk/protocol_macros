Attribute VB_Name = "make_approval"
Option Explicit

' Собирает лист согласования (шаблон 0.3)
Function BuildApprovalPage(wbMain As Workbook, wdApp As Object, Optional savePath As String = "") As Object

    Dim approvalPath As String
    Dim protNum As String
    Dim date1 As String
    Dim wdApproval As Object

    approvalPath = GetSourcePathByNumber_Common(wbMain, "0.3")

    If approvalPath = "" Then
        MsgBox "Не найден путь к шаблону approval (0.3) в Sources."
        Exit Function
    End If

    Set wdApproval = wdApp.Documents.Add(Template:=approvalPath)

    protNum = GetConfigValue(wbMain, "{{ПРОТ_НОМЕР}}")
    date1 = GetConfigValue(wbMain, "{{ДАТА1}}")

    Call ReplacePlaceholderInWord(wdApproval, "{{ПРОТ_НОМЕР}}", protNum)
    Call ReplacePlaceholderInWord(wdApproval, "{{ДАТА1}}", date1)

    If savePath <> "" Then
        wdApproval.SaveAs2 Filename:=savePath
    End If

    Set BuildApprovalPage = wdApproval

End Function


Sub RunApprovalOnly()

    Dim wbMain As Workbook
    Dim wdApp As Object
    Dim wdApproval As Object

    Set wbMain = ThisWorkbook

    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = True

    Set wdApproval = BuildApprovalPage(wbMain, wdApp)

End Sub
