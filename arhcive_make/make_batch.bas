Attribute VB_Name = "make_batch"
Option Explicit

' Главный батч: собирает многоадресный протокол воедино.
' Структура:
' 1) титульник
' 2) тела по Номерам 1..N из Sources
' 3) лист согласования
Sub BuildMultiAddressProtocol()

    Dim wbMain As Workbook
    Dim wsSources As Worksheet
    Dim wdApp As Object
    Dim wdMaster As Object

    Dim lastRow As Long
    Dim i As Long
    Dim srcNum As Variant

    Dim titleFile As String
    Dim bodyFile As String
    Dim approvalFile As String

    Set wbMain = ThisWorkbook
    Set wsSources = wbMain.Worksheets("Sources")

    ' Создаём единый Word-экземпляр и пустой мастер-документ
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = True
    Set wdMaster = wdApp.Documents.Add

    ' 1. Титульник
    titleFile = MakeTitleFile(wbMain, wdApp)   ' функция в модуле make_title
    If titleFile <> "" Then
        With wdMaster.Range
            .Collapse 0
            .InsertFile FileName:=titleFile
        End With
    End If

    ' 2. Тела по многим адресам (Номер = 1..N в Sources)
    lastRow = wsSources.Cells(wsSources.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow

        srcNum = wsSources.Cells(i, 1).Value   ' колонка "Номер"[file:343]

        ' Берём только объектные Номера: 1,2,3,... (0.1...0.6 – это шаблоны)
        If IsNumeric(srcNum) Then
            If CLng(srcNum) >= 1 Then

                ' Тело для этого Номера
                bodyFile = MakeBodyFileForSource(wbMain, wdApp, CLng(srcNum))

                If bodyFile <> "" Then
                    ' Вставляем тело в конец мастер-документа
                    With wdMaster.Range
                        .Collapse 0
                        .InsertFile FileName:=bodyFile
                    End With
                End If

            End If
        End If
    Next i

    ' 3. Лист согласования (один на весь протокол)
    approvalFile = MakeApprovalFile(wbMain, wdApp)   ' функция в модуле make_approval
    If approvalFile <> "" Then
        With wdMaster.Range
            .Collapse 0
            .InsertFile FileName:=approvalFile
        End With
    End If

    ' 4. Сохраняем мастер
    wdMaster.SaveAs2 Filename:=wbMain.Path & "\Protocol_multi_address.docx"

    MsgBox "Многоадресный протокол собран: " & wbMain.Path & "\Protocol_multi_address.docx"

End Sub
