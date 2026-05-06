Attribute VB_Name = "Sap2000ExcelMvp"
Option Explicit

' =============================================================================
' SAP2000 Excel VBA MVP - criação automática de pórtico simples
' =============================================================================
' Objetivo:
'   - Ler os dados digitados na aba "Entrada".
'   - Criar um modelo NOVO do zero no SAP2000 via OAPI/COM.
'   - Lançar dois pilares, uma viga superior, apoios engastados e cargas básicas.
'
' Observações:
'   - O código usa late binding (Object/CreateObject/GetObject), portanto não exige
'     referência manual à biblioteca do SAP2000 no Editor VBA.
'   - As chamadas OAPI dependem do SAP2000 instalado e registrado no Windows.
'   - Unidades adotadas no SAP2000: tf, m, C (enum Ton_m_C = 12).
' =============================================================================

Private Const SHEET_ENTRADA As String = "Entrada"
Private Const SHEET_LOG As String = "Log"

Private Const SAP_UNITS_TF_M_C As Long = 12
Private Const MAT_TYPE_CONCRETE As Long = 2
Private Const LOAD_DEAD As Long = 1
Private Const LOAD_LIVE As Long = 3
Private Const LOAD_WIND As Long = 6
Private Const ITEMTYPE_OBJECTS As Long = 0
Private Const DIST_LOAD_FORCE_LENGTH As Long = 1
Private Const DIR_GLOBAL_Z As Long = 6

Private Const CONCRETE_UNIT_WEIGHT_TF_M3 As Double = 2.5
Private Const MPA_TO_TF_M2 As Double = 101.9716213

Private SapObject As Object
Private SapModel As Object

' -----------------------------------------------------------------------------
' Cria a aba de entrada, a aba de log, os campos padrão e os botões de comando.
' Esta macro é chamada pelo script CriarPlanilhaSAP2000.vbs e também pode ser
' executada manualmente a qualquer momento para reconstruir o layout.
' -----------------------------------------------------------------------------
Public Sub CriarLayout()
    On Error GoTo TrataErro

    Dim ws As Worksheet

    Application.ScreenUpdating = False

    Set ws = EnsureSheet(SHEET_ENTRADA)
    ws.Cells.Clear
    RemoverBotoes ws

    ws.Range("A1").Value = "SAP2000 - Pórtico simples MVP"
    ws.Range("A1:D1").Merge
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 16

    ws.Range("A3").Value = "Dados gerais"
    ws.Range("A3").Font.Bold = True
    EscreverCampo ws, 4, "Vão do pórtico [cm]", 600#
    EscreverCampo ws, 5, "Pé-direito [cm]", 300#

    ws.Range("A7").Value = "Dados da viga"
    ws.Range("A7").Font.Bold = True
    EscreverCampo ws, 8, "Largura da viga [cm]", 20#
    EscreverCampo ws, 9, "Altura da viga [cm]", 50#
    EscreverCampo ws, 10, "Fck da viga [MPa]", 25#

    ws.Range("A12").Value = "Dados dos pilares"
    ws.Range("A12").Font.Bold = True
    EscreverCampo ws, 13, "Dimensão X do pilar [cm]", 25#
    EscreverCampo ws, 14, "Dimensão Y do pilar [cm]", 25#
    EscreverCampo ws, 15, "Fck dos pilares [MPa]", 25#

    ws.Range("A17").Value = "Carregamentos"
    ws.Range("A17").Font.Bold = True
    EscreverCampo ws, 18, "Carga distribuída vertical na viga [tf/m]", 1#
    EscreverCampo ws, 19, "Carga horizontal de vento [tf]", 2#

    ws.Range("D3").Value = "Botões"
    ws.Range("D3").Font.Bold = True
    AdicionarBotao ws, "Criar Layout", "CriarLayout", 4
    AdicionarBotao ws, "Exportar para SAP2000", "ExportarParaSAP2000", 7
    AdicionarBotao ws, "Limpar Log", "LimparLog", 10

    ws.Range("A22").Value = "Notas"
    ws.Range("A22").Font.Bold = True
    ws.Range("A23").Value = "1. Preencha os valores em centímetros, MPa, tf/m e tf."
    ws.Range("A24").Value = "2. Clique em Exportar para SAP2000 para criar um modelo novo do zero."
    ws.Range("A25").Value = "3. O SAP2000 deve estar instalado no Windows para a exportação real funcionar."

    ws.Columns("A:D").AutoFit
    ws.Range("B4:B19").Interior.Color = RGB(255, 255, 204)
    ws.Activate

    PrepararLog
    RegistrarLog "INFO", "Layout criado/atualizado."

Finalizar:
    Application.ScreenUpdating = True
    Exit Sub

TrataErro:
    Application.ScreenUpdating = True
    MsgBox "Erro ao criar o layout: " & Err.Description, vbCritical, "SAP2000 MVP"
End Sub

' -----------------------------------------------------------------------------
' Macro principal: valida os dados da planilha e cria um pórtico novo no SAP2000.
' Geometria:
'   Pilar esquerdo: (0,0,0) até (0,0,H)
'   Pilar direito:  (L,0,0) até (L,0,H)
'   Viga superior:  (0,0,H) até (L,0,H)
' -----------------------------------------------------------------------------
Public Sub ExportarParaSAP2000()
    On Error GoTo TrataErro

    Dim ws As Worksheet
    Dim vaoCm As Double, peDireitoCm As Double
    Dim larguraVigaCm As Double, alturaVigaCm As Double, fckVigaMPa As Double
    Dim dimPilarXCm As Double, dimPilarYCm As Double, fckPilarMPa As Double
    Dim cargaVigaTfM As Double, cargaVentoTf As Double
    Dim L As Double, H As Double
    Dim larguraVigaM As Double, alturaVigaM As Double
    Dim dimPilarXM As Double, dimPilarYM As Double

    Dim matViga As String, matPilar As String
    Dim secViga As String, secPilar As String
    Dim framePilarEsq As String, framePilarDir As String, frameViga As String
    Dim noBaseEsq As String, noTopoEsq As String, noBaseDir As String, noTopoDir As String

    Set ws = EnsureSheet(SHEET_ENTRADA)
    PrepararLog
    RegistrarLog "INFO", "Iniciando exportação para SAP2000."

    vaoCm = LerNumeroPositivo(ws, "B4", "Vão do pórtico [cm]")
    peDireitoCm = LerNumeroPositivo(ws, "B5", "Pé-direito [cm]")
    larguraVigaCm = LerNumeroPositivo(ws, "B8", "Largura da viga [cm]")
    alturaVigaCm = LerNumeroPositivo(ws, "B9", "Altura da viga [cm]")
    fckVigaMPa = LerNumeroPositivo(ws, "B10", "Fck da viga [MPa]")
    dimPilarXCm = LerNumeroPositivo(ws, "B13", "Dimensão X do pilar [cm]")
    dimPilarYCm = LerNumeroPositivo(ws, "B14", "Dimensão Y do pilar [cm]")
    fckPilarMPa = LerNumeroPositivo(ws, "B15", "Fck dos pilares [MPa]")
    cargaVigaTfM = LerNumeroPositivo(ws, "B18", "Carga distribuída vertical na viga [tf/m]")
    cargaVentoTf = LerNumeroPositivo(ws, "B19", "Carga horizontal de vento [tf]")
    RegistrarLog "INFO", "Dados de entrada validados."

    L = vaoCm / 100#
    H = peDireitoCm / 100#
    larguraVigaM = larguraVigaCm / 100#
    alturaVigaM = alturaVigaCm / 100#
    dimPilarXM = dimPilarXCm / 100#
    dimPilarYM = dimPilarYCm / 100#
    RegistrarLog "INFO", "Conversão cm -> m concluída: L=" & Format(L, "0.000") & " m; H=" & Format(H, "0.000") & " m."

    ConectarOuAbrirSAP2000
    CriarModeloNovo

    matViga = "CONC_VIGA_FCK" & LimparNome(CStr(fckVigaMPa))
    matPilar = "CONC_PILAR_FCK" & LimparNome(CStr(fckPilarMPa))
    CriarMaterialConcreto matViga, fckVigaMPa
    CriarMaterialConcreto matPilar, fckPilarMPa

    secPilar = "PILAR_RET_" & LimparNome(CStr(dimPilarXCm)) & "x" & LimparNome(CStr(dimPilarYCm)) & "cm"
    secViga = "VIGA_RET_" & LimparNome(CStr(larguraVigaCm)) & "x" & LimparNome(CStr(alturaVigaCm)) & "cm"

    ' SetRectangle usa dimensões na unidade atual do SAP2000 (m).
    ' Para o pilar, t3 recebe a dimensão X e t2 recebe a dimensão Y.
    CheckSapRet SapModel.PropFrame.SetRectangle(secPilar, matPilar, dimPilarXM, dimPilarYM), "Criar seção retangular dos pilares"
    CheckSapRet SapModel.PropFrame.SetRectangle(secViga, matViga, alturaVigaM, larguraVigaM), "Criar seção retangular da viga"
    RegistrarLog "INFO", "Seções retangulares criadas."

    CheckSapRet SapModel.FrameObj.AddByCoord(0#, 0#, 0#, 0#, 0#, H, framePilarEsq, secPilar, "Pilar_Esquerdo", "Global"), "Criar pilar esquerdo"
    CheckSapRet SapModel.FrameObj.AddByCoord(L, 0#, 0#, L, 0#, H, framePilarDir, secPilar, "Pilar_Direito", "Global"), "Criar pilar direito"
    CheckSapRet SapModel.FrameObj.AddByCoord(0#, 0#, H, L, 0#, H, frameViga, secViga, "Viga_Superior", "Global"), "Criar viga superior"
    RegistrarLog "INFO", "Elementos de barra criados: dois pilares e uma viga superior."

    CheckSapRet SapModel.FrameObj.GetPoints(framePilarEsq, noBaseEsq, noTopoEsq), "Obter nós do pilar esquerdo"
    CheckSapRet SapModel.FrameObj.GetPoints(framePilarDir, noBaseDir, noTopoDir), "Obter nós do pilar direito"

    AplicarEngaste noBaseEsq
    AplicarEngaste noBaseDir
    RegistrarLog "INFO", "Engastes aplicados nas bases dos pilares."

    CriarPadroesECasosDeCarga
    AplicarCargaDistribuidaVertical frameViga, cargaVigaTfM
    AplicarCargaHorizontal noTopoDir, cargaVentoTf

    CheckSapRet SapModel.View.RefreshView(0, False), "Atualizar visualização do SAP2000"
    RegistrarLog "INFO", "Modelo do pórtico criado com sucesso no SAP2000."
    MsgBox "Modelo criado no SAP2000 com sucesso.", vbInformation, "SAP2000 MVP"
    Exit Sub

TrataErro:
    RegistrarLog "ERRO", Err.Description
    MsgBox "Não foi possível exportar para o SAP2000." & vbCrLf & vbCrLf & Err.Description, vbCritical, "SAP2000 MVP"
End Sub

' -----------------------------------------------------------------------------
' Limpa a aba de log e recria o cabeçalho.
' -----------------------------------------------------------------------------
Public Sub LimparLog()
    On Error GoTo TrataErro
    EnsureSheet(SHEET_LOG).Cells.Clear
    PrepararLog
    RegistrarLog "INFO", "Log limpo."
    Exit Sub

TrataErro:
    MsgBox "Erro ao limpar o log: " & Err.Description, vbCritical, "SAP2000 MVP"
End Sub

Private Sub ConectarOuAbrirSAP2000()
    On Error Resume Next
    Set SapObject = GetObject(, "CSI.SAP2000.API.SapObject")
    On Error GoTo 0

    If SapObject Is Nothing Then
        RegistrarLog "INFO", "Nenhuma instância aberta encontrada. Iniciando SAP2000."
        Set SapObject = CreateObject("CSI.SAP2000.API.SapObject")
        SapObject.ApplicationStart SAP_UNITS_TF_M_C, True
    Else
        RegistrarLog "INFO", "Conectado a uma instância existente do SAP2000."
    End If

    Set SapModel = SapObject.SapModel
    If SapModel Is Nothing Then
        Err.Raise vbObjectError + 3000, "ConectarOuAbrirSAP2000", "Não foi possível obter SapModel pela OAPI."
    End If
End Sub

Private Sub CriarModeloNovo()
    CheckSapRet SapModel.InitializeNewModel(SAP_UNITS_TF_M_C), "Inicializar modelo novo com unidades tf, m, C"
    CheckSapRet SapModel.File.NewBlank, "Criar arquivo/modelo em branco"
    CheckSapRet SapModel.SetPresentUnits(SAP_UNITS_TF_M_C), "Definir unidades atuais em tf, m, C"
    RegistrarLog "INFO", "Modelo novo em branco criado no SAP2000."
End Sub

Private Sub CriarMaterialConcreto(ByVal nomeMaterial As String, ByVal fckMPa As Double)
    Dim moduloElasticidadeTfM2 As Double

    ' Estimativa usual para concreto: Eci ~= 5600 * sqrt(fck) em MPa.
    ' Convertida para tf/m² para compatibilizar com tf, m, C.
    moduloElasticidadeTfM2 = 5600# * Sqr(fckMPa) * MPA_TO_TF_M2

    CheckSapRet SapModel.PropMaterial.SetMaterial(nomeMaterial, MAT_TYPE_CONCRETE), "Criar material " & nomeMaterial
    CheckSapRet SapModel.PropMaterial.SetMPIsotropic(nomeMaterial, moduloElasticidadeTfM2, 0.2, 0.0000099), "Definir propriedades isotrópicas de " & nomeMaterial
    CheckSapRet SapModel.PropMaterial.SetWeightAndMass(nomeMaterial, 1, CONCRETE_UNIT_WEIGHT_TF_M3), "Definir peso específico de " & nomeMaterial
    RegistrarLog "INFO", "Material criado: " & nomeMaterial & " (fck=" & Format(fckMPa, "0.###") & " MPa)."
End Sub

Private Sub CriarPadroesECasosDeCarga()
    CriarPadraoECasoEstatico "PP", LOAD_DEAD, 1#
    CriarPadraoECasoEstatico "SOBRECARGA", LOAD_LIVE, 0#
    CriarPadraoECasoEstatico "VENTO", LOAD_WIND, 0#
    RegistrarLog "INFO", "Casos de carga criados: PP, SOBRECARGA e VENTO."
End Sub

Private Sub CriarPadraoECasoEstatico(ByVal nome As String, ByVal tipo As Long, ByVal multiplicadorPesoProprio As Double)
    Dim loadTypes(0 To 0) As String
    Dim loadNames(0 To 0) As String
    Dim scaleFactors(0 To 0) As Double
    Dim ret As Long

    ret = SapModel.LoadPatterns.Add(nome, tipo, multiplicadorPesoProprio, True)
    If ret <> 0 Then
        RegistrarLog "AVISO", "Load pattern '" & nome & "' retornou código " & CStr(ret) & ". Tentando seguir."
    End If

    loadTypes(0) = "Load"
    loadNames(0) = nome
    scaleFactors(0) = 1#

    CheckSapRet SapModel.LoadCases.StaticLinear.SetCase(nome), "Criar caso estático linear " & nome
    CheckSapRet SapModel.LoadCases.StaticLinear.SetLoads(nome, 1, loadTypes, loadNames, scaleFactors), "Associar padrão ao caso " & nome
End Sub

Private Sub AplicarEngaste(ByVal nomeNo As String)
    Dim restricoes(0 To 5) As Boolean
    Dim i As Long

    For i = 0 To 5
        restricoes(i) = True
    Next i

    CheckSapRet SapModel.PointObj.SetRestraint(nomeNo, restricoes), "Aplicar engaste no nó " & nomeNo
End Sub

Private Sub AplicarCargaDistribuidaVertical(ByVal nomeFrame As String, ByVal cargaTfM As Double)
    ' Carga vertical global para baixo: direção Global Z com sinal negativo.
    CheckSapRet SapModel.FrameObj.SetLoadDistributed(nomeFrame, "SOBRECARGA", DIST_LOAD_FORCE_LENGTH, DIR_GLOBAL_Z, 0#, 1#, -cargaTfM, -cargaTfM, "Global", True, True, ITEMTYPE_OBJECTS), "Aplicar carga distribuída vertical na viga"
    RegistrarLog "INFO", "Carga distribuída aplicada na viga: " & Format(cargaTfM, "0.###") & " tf/m para baixo."
End Sub

Private Sub AplicarCargaHorizontal(ByVal nomeNo As String, ByVal cargaTf As Double)
    Dim cargas(0 To 5) As Double

    ' Força horizontal no nó superior direito no sentido +X global.
    cargas(0) = cargaTf
    cargas(1) = 0#
    cargas(2) = 0#
    cargas(3) = 0#
    cargas(4) = 0#
    cargas(5) = 0#

    CheckSapRet SapModel.PointObj.SetLoadForce(nomeNo, "VENTO", cargas, True, "Global", ITEMTYPE_OBJECTS), "Aplicar carga horizontal de vento"
    RegistrarLog "INFO", "Carga horizontal de vento aplicada no nó superior direito: " & Format(cargaTf, "0.###") & " tf em +X."
End Sub

Private Function LerNumeroPositivo(ByVal ws As Worksheet, ByVal endereco As String, ByVal nomeCampo As String) As Double
    Dim valor As Variant

    valor = ws.Range(endereco).Value
    If Not IsNumeric(valor) Then
        Err.Raise vbObjectError + 3100, "Validação", "Campo inválido: " & nomeCampo & " deve ser numérico."
    End If

    LerNumeroPositivo = CDbl(valor)
    If LerNumeroPositivo <= 0# Then
        Err.Raise vbObjectError + 3101, "Validação", "Campo inválido: " & nomeCampo & " deve ser maior que zero."
    End If
End Function

Private Sub EscreverCampo(ByVal ws As Worksheet, ByVal linha As Long, ByVal rotulo As String, ByVal valorPadrao As Double)
    ws.Cells(linha, 1).Value = rotulo
    ws.Cells(linha, 2).Value = valorPadrao
    ws.Cells(linha, 3).Value = "Entrada do usuário"
End Sub

Private Sub AdicionarBotao(ByVal ws As Worksheet, ByVal texto As String, ByVal macroNome As String, ByVal linha As Long)
    Dim btn As Button
    Set btn = ws.Buttons.Add(ws.Cells(linha, 4).Left, ws.Cells(linha, 4).Top, 150, 28)
    btn.Caption = texto
    btn.OnAction = macroNome
End Sub

Private Sub RemoverBotoes(ByVal ws As Worksheet)
    Dim shp As Shape
    For Each shp In ws.Shapes
        If shp.Type = msoFormControl Then
            shp.Delete
        End If
    Next shp
End Sub

Private Function EnsureSheet(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set EnsureSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If EnsureSheet Is Nothing Then
        Set EnsureSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        EnsureSheet.Name = sheetName
    End If
End Function

Private Sub PrepararLog()
    Dim ws As Worksheet
    Set ws = EnsureSheet(SHEET_LOG)

    If Len(CStr(ws.Range("A1").Value)) = 0 Then
        ws.Range("A1:C1").Value = Array("Data/Hora", "Nível", "Mensagem")
        ws.Range("A1:C1").Font.Bold = True
        ws.Columns("A:C").AutoFit
    End If
End Sub

Private Sub RegistrarLog(ByVal nivel As String, ByVal mensagem As String)
    Dim ws As Worksheet
    Dim linha As Long

    Set ws = EnsureSheet(SHEET_LOG)
    If Len(CStr(ws.Range("A1").Value)) = 0 Then PrepararLog

    linha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(linha, 1).Value = Now
    ws.Cells(linha, 2).Value = nivel
    ws.Cells(linha, 3).Value = mensagem
    ws.Columns("A:C").AutoFit
End Sub

Private Sub CheckSapRet(ByVal ret As Long, ByVal acao As String)
    If ret <> 0 Then
        RegistrarLog "ERRO", acao & " falhou. Código SAP2000: " & CStr(ret)
        Err.Raise vbObjectError + 3200, "SAP2000 OAPI", acao & " falhou. Código SAP2000: " & CStr(ret)
    End If
End Sub

Private Function LimparNome(ByVal texto As String) As String
    LimparNome = Replace(Replace(Replace(texto, ",", "_"), ".", "_"), " ", "_")
End Function
