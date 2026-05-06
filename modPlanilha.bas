Attribute VB_Name = "modPlanilha"
Option Explicit

' Módulo responsável por montar a aba Entrada e ler os dados digitados.

Private Const COL_LABEL As Long = 1
Private Const COL_VALUE As Long = 2
Private Const COL_UNIT As Long = 3

Public Const ROW_VAO As Long = 4
Public Const ROW_PE_DIREITO As Long = 5
Public Const ROW_LARGURA_VIGA As Long = 8
Public Const ROW_ALTURA_VIGA As Long = 9
Public Const ROW_FCK_VIGA As Long = 10
Public Const ROW_PILAR_DIM_X As Long = 13
Public Const ROW_PILAR_DIM_Y As Long = 14
Public Const ROW_FCK_PILAR As Long = 15
Public Const ROW_CARGA_DISTRIBUIDA As Long = 18
Public Const ROW_CARGA_VENTO As Long = 19

Public Sub CriarOuAtualizarLayout()
    On Error GoTo TratarErro

    Dim ws As Worksheet
    Set ws = EnsureEntradaSheet()

    Application.ScreenUpdating = False

    ws.Cells.Clear
    RemoverBotoes ws

    ws.Range("A1").Value = "MVP Excel → SAP2000: Pórtico plano simples"
    ws.Range("A1:C1").Merge
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14

    CriarSecao ws, 3, "Dados gerais"
    DefinirCampo ws, ROW_VAO, "Vão/comprimento do pórtico", "cm", 600
    DefinirCampo ws, ROW_PE_DIREITO, "Pé-direito", "cm", 300

    CriarSecao ws, 7, "Dados da viga"
    DefinirCampo ws, ROW_LARGURA_VIGA, "Largura da viga", "cm", 20
    DefinirCampo ws, ROW_ALTURA_VIGA, "Altura da viga", "cm", 50
    DefinirCampo ws, ROW_FCK_VIGA, "Fck da viga", "MPa", 30

    CriarSecao ws, 12, "Dados dos pilares"
    DefinirCampo ws, ROW_PILAR_DIM_X, "Dimensão X do pilar", "cm", 30
    DefinirCampo ws, ROW_PILAR_DIM_Y, "Dimensão Y do pilar", "cm", 30
    DefinirCampo ws, ROW_FCK_PILAR, "Fck dos pilares", "MPa", 30

    CriarSecao ws, 17, "Carregamentos"
    DefinirCampo ws, ROW_CARGA_DISTRIBUIDA, "Carga distribuída vertical na viga", "tf/m", 1
    DefinirCampo ws, ROW_CARGA_VENTO, "Carga horizontal de vento", "tf", 2

    CriarBotoes ws
    FormatarEntrada ws
    EnsureLogSheet

    Application.ScreenUpdating = True
    RegistrarLog "INFO", "Layout criado/atualizado."
    MsgBox "Layout criado/atualizado com sucesso.", vbInformation, "SAP2000 MVP"
    Exit Sub

TratarErro:
    Application.ScreenUpdating = True
    RegistrarLog "ERRO", "Falha ao criar layout: " & Err.Description
    MsgBox "Não foi possível criar/atualizar o layout: " & Err.Description, vbCritical, "SAP2000 MVP"
End Sub

Public Function LerDadosEntrada(ByRef dados As DadosPortico) As Boolean
    On Error GoTo TratarErro

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(ENTRADA_SHEET_NAME)

    dados.VaoCm = LerNumeroObrigatorio(ws, ROW_VAO, "Vão/comprimento do pórtico [cm]")
    dados.PeDireitoCm = LerNumeroObrigatorio(ws, ROW_PE_DIREITO, "Pé-direito [cm]")
    dados.LarguraVigaCm = LerNumeroObrigatorio(ws, ROW_LARGURA_VIGA, "Largura da viga [cm]")
    dados.AlturaVigaCm = LerNumeroObrigatorio(ws, ROW_ALTURA_VIGA, "Altura da viga [cm]")
    dados.FckVigaMPa = LerNumeroObrigatorio(ws, ROW_FCK_VIGA, "Fck da viga [MPa]")
    dados.PilarDimXCm = LerNumeroObrigatorio(ws, ROW_PILAR_DIM_X, "Dimensão X do pilar [cm]")
    dados.PilarDimYCm = LerNumeroObrigatorio(ws, ROW_PILAR_DIM_Y, "Dimensão Y do pilar [cm]")
    dados.FckPilarMPa = LerNumeroObrigatorio(ws, ROW_FCK_PILAR, "Fck dos pilares [MPa]")
    dados.CargaDistribuidaTfM = LerNumeroObrigatorio(ws, ROW_CARGA_DISTRIBUIDA, "Carga distribuída vertical na viga [tf/m]")
    dados.CargaVentoTf = LerNumeroObrigatorio(ws, ROW_CARGA_VENTO, "Carga horizontal de vento [tf]")

    ConverterDadosParaUnidadesSAP dados
    LerDadosEntrada = True
    Exit Function

TratarErro:
    RegistrarLog "ERRO", "Erro ao ler a aba Entrada: " & Err.Description
    LerDadosEntrada = False
End Function

Public Function EnsureEntradaSheet() As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(ENTRADA_SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        ws.Name = ENTRADA_SHEET_NAME
    End If

    Set EnsureEntradaSheet = ws
End Function

Private Function LerNumeroObrigatorio(ByVal ws As Worksheet, ByVal rowNumber As Long, ByVal nomeCampo As String) As Double
    Dim valor As Variant

    valor = ws.Cells(rowNumber, COL_VALUE).Value

    If Len(Trim$(CStr(valor))) = 0 Then
        Err.Raise vbObjectError + 100, "Entrada", "O campo '" & nomeCampo & "' está vazio."
    End If

    If Not IsNumeric(valor) Then
        Err.Raise vbObjectError + 101, "Entrada", "O campo '" & nomeCampo & "' deve conter um número."
    End If

    LerNumeroObrigatorio = CDbl(valor)
End Function

Private Sub CriarSecao(ByVal ws As Worksheet, ByVal rowNumber As Long, ByVal titulo As String)
    ws.Cells(rowNumber, COL_LABEL).Value = titulo
    ws.Range(ws.Cells(rowNumber, COL_LABEL), ws.Cells(rowNumber, COL_UNIT)).Merge
    With ws.Cells(rowNumber, COL_LABEL)
        .Font.Bold = True
        .Interior.Color = RGB(217, 225, 242)
    End With
End Sub

Private Sub DefinirCampo(ByVal ws As Worksheet, ByVal rowNumber As Long, ByVal rotulo As String, ByVal unidade As String, ByVal valorPadrao As Double)
    ws.Cells(rowNumber, COL_LABEL).Value = rotulo
    ws.Cells(rowNumber, COL_VALUE).Value = valorPadrao
    ws.Cells(rowNumber, COL_UNIT).Value = unidade
End Sub

Private Sub FormatarEntrada(ByVal ws As Worksheet)
    ws.Columns("A:A").ColumnWidth = 42
    ws.Columns("B:B").ColumnWidth = 16
    ws.Columns("C:C").ColumnWidth = 12
    ws.Range("B4:B19").NumberFormat = "0.00"
    ws.Range("A:C").VerticalAlignment = xlCenter
    ws.Range("A1:C19").Borders.LineStyle = xlContinuous
    ws.Activate
    ws.Range("B4").Select
End Sub

Private Sub CriarBotoes(ByVal ws As Worksheet)
    AdicionarBotao ws, "btnLayout", "Criar/Atualizar Layout", "CriarOuAtualizarLayout", 5, 330, 170, 28
    AdicionarBotao ws, "btnExportar", "Exportar para SAP2000", "ExportarParaSAP2000", 185, 330, 170, 28
    AdicionarBotao ws, "btnLimparLog", "Limpar Log", "LimparLog", 365, 330, 120, 28
End Sub

Private Sub AdicionarBotao(ByVal ws As Worksheet, ByVal nome As String, ByVal texto As String, ByVal macro As String, ByVal leftPos As Double, ByVal topPos As Double, ByVal widthSize As Double, ByVal heightSize As Double)
    Dim btn As Button
    Set btn = ws.Buttons.Add(leftPos, topPos, widthSize, heightSize)
    btn.Name = nome
    btn.Caption = texto
    btn.OnAction = macro
End Sub

Private Sub RemoverBotoes(ByVal ws As Worksheet)
    Dim shp As Shape

    For Each shp In ws.Shapes
        If Left$(shp.Name, 3) = "btn" Then shp.Delete
    Next shp
End Sub
