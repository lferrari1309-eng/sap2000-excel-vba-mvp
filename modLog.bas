Attribute VB_Name = "modLog"
Option Explicit

' Módulo responsável por registrar mensagens de execução na aba Log.

Public Const LOG_SHEET_NAME As String = "Log"

Private Const LOG_HEADER_DATE As String = "Data/Hora"
Private Const LOG_HEADER_LEVEL As String = "Nível"
Private Const LOG_HEADER_MESSAGE As String = "Mensagem"

Public Sub LimparLog()
    Dim ws As Worksheet
    Set ws = EnsureLogSheet()

    ws.Cells.Clear
    PrepararCabecalhoLog ws
    RegistrarLog "INFO", "Log limpo."
End Sub

Public Sub RegistrarLog(ByVal nivel As String, ByVal mensagem As String)
    Dim ws As Worksheet
    Dim nextRow As Long

    Set ws = EnsureLogSheet()
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    If nextRow < 2 Then nextRow = 2

    ws.Cells(nextRow, 1).Value = Now
    ws.Cells(nextRow, 2).Value = nivel
    ws.Cells(nextRow, 3).Value = mensagem
    ws.Columns("A:C").AutoFit
End Sub

Public Function EnsureLogSheet() As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(LOG_SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = LOG_SHEET_NAME
        PrepararCabecalhoLog ws
    ElseIf Len(CStr(ws.Cells(1, 1).Value)) = 0 Then
        PrepararCabecalhoLog ws
    End If

    Set EnsureLogSheet = ws
End Function

Private Sub PrepararCabecalhoLog(ByVal ws As Worksheet)
    ws.Range("A1").Value = LOG_HEADER_DATE
    ws.Range("B1").Value = LOG_HEADER_LEVEL
    ws.Range("C1").Value = LOG_HEADER_MESSAGE

    With ws.Range("A1:C1")
        .Font.Bold = True
        .Interior.Color = RGB(220, 230, 241)
    End With

    ws.Columns("A:A").NumberFormat = "dd/mm/yyyy hh:mm:ss"
    ws.Columns("A:C").AutoFit
End Sub
