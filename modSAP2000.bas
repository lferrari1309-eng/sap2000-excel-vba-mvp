Attribute VB_Name = "modSAP2000"
Option Explicit

' Módulo de integração com o SAP2000 via OAPI usando late binding.
' Late binding evita marcar uma referência COM fixa no VBA e reduz conflitos entre versões.

Private Const SAP_PROG_ID As String = "CSI.SAP2000.API.SapObject"
Private Const SAP_HELPER_PROG_ID As String = "SAP2000v1.Helper"

' Enumerações numéricas da OAPI. Podem variar em versões muito antigas do SAP2000.
Private Const SAP_UNITS_TON_M_C As Long = 12
Private Const SAP_MATERIAL_CONCRETE As Long = 2
Private Const SAP_LOADTYPE_DEAD As Long = 1
Private Const SAP_LOADTYPE_LIVE As Long = 3
Private Const SAP_LOADTYPE_WIND As Long = 6
Private Const SAP_DIST_LOAD_FORCE_PER_LENGTH As Long = 1
Private Const SAP_DIR_GLOBAL_X As Long = 4
Private Const SAP_DIR_GLOBAL_Z As Long = 6
Private Const SAP_COORD_GLOBAL As String = "Global"

Private Const CONCRETE_POISSON As Double = 0.2
Private Const CONCRETE_THERMAL_COEFF As Double = 0.0000099

Public Function ConectarOuAbrirSAP2000(ByRef sapObject As Object, ByRef sapModel As Object) As Boolean
    On Error GoTo TratarErro

    Dim helper As Object

    RegistrarLog "INFO", "Tentando conectar a uma instância aberta do SAP2000."

    On Error Resume Next
    Set sapObject = GetObject(, SAP_PROG_ID)
    On Error GoTo TratarErro

    If sapObject Is Nothing Then
        RegistrarLog "INFO", "Nenhuma instância encontrada. Abrindo SAP2000 via OAPI."
        Set helper = CreateObject(SAP_HELPER_PROG_ID)
        Set sapObject = helper.CreateObjectProgID(SAP_PROG_ID)
        sapObject.ApplicationStart
    Else
        RegistrarLog "INFO", "Instância aberta do SAP2000 encontrada."
    End If

    Set sapModel = sapObject.SapModel
    ConectarOuAbrirSAP2000 = Not sapModel Is Nothing
    Exit Function

TratarErro:
    RegistrarLog "ERRO", "Falha ao conectar/abrir SAP2000: " & Err.Description
    ConectarOuAbrirSAP2000 = False
End Function

Public Function CriarModeloPorticoSAP2000(ByVal sapModel As Object, ByRef dados As DadosPortico) As Boolean
    On Error GoTo TratarErro

    Dim ret As Long
    Dim pBaseEsq As String
    Dim pTopoEsq As String
    Dim pBaseDir As String
    Dim pTopoDir As String
    Dim framePilarEsq As String
    Dim framePilarDir As String
    Dim frameViga As String

    RegistrarLog "INFO", "Inicializando novo modelo em tf, m, C."
    ret = sapModel.InitializeNewModel(SAP_UNITS_TON_M_C)
    VerificarRetorno ret, "InitializeNewModel"

    ret = sapModel.File.NewBlank
    VerificarRetorno ret, "File.NewBlank"

    ret = sapModel.SetPresentUnits(SAP_UNITS_TON_M_C)
    VerificarRetorno ret, "SetPresentUnits"

    DefinirMateriais sapModel, dados
    DefinirSecoes sapModel, dados
    DefinirPadroesCarga sapModel

    RegistrarLog "INFO", "Criando barras do pórtico."
    framePilarEsq = AddFrameByCoord(sapModel, 0, 0, 0, 0, 0, dados.PeDireitoM, FRAME_PILAR_ESQUERDO, SECAO_PILAR)
    framePilarDir = AddFrameByCoord(sapModel, dados.VaoM, 0, 0, dados.VaoM, 0, dados.PeDireitoM, FRAME_PILAR_DIREITO, SECAO_PILAR)
    frameViga = AddFrameByCoord(sapModel, 0, 0, dados.PeDireitoM, dados.VaoM, 0, dados.PeDireitoM, FRAME_VIGA_SUPERIOR, SECAO_VIGA)

    ObterPontosFrame sapModel, framePilarEsq, pBaseEsq, pTopoEsq
    ObterPontosFrame sapModel, framePilarDir, pBaseDir, pTopoDir

    AplicarEngaste sapModel, pBaseEsq
    AplicarEngaste sapModel, pBaseDir
    AplicarCargaDistribuidaViga sapModel, frameViga, dados.CargaDistribuidaTfM
    AplicarCargaVento sapModel, pTopoDir, dados.CargaVentoTf

    ret = sapModel.View.RefreshView(0, False)
    VerificarRetorno ret, "View.RefreshView"

    RegistrarLog "INFO", "Modelo do pórtico criado com sucesso no SAP2000."
    CriarModeloPorticoSAP2000 = True
    Exit Function

TratarErro:
    RegistrarLog "ERRO", "Falha ao criar modelo no SAP2000: " & Err.Description
    CriarModeloPorticoSAP2000 = False
End Function

Private Sub DefinirMateriais(ByVal sapModel As Object, ByRef dados As DadosPortico)
    RegistrarLog "INFO", "Definindo materiais de concreto."
    DefinirMaterialConcreto sapModel, MATERIAL_CONCRETO_VIGA, dados.FckVigaMPa
    DefinirMaterialConcreto sapModel, MATERIAL_CONCRETO_PILAR, dados.FckPilarMPa
End Sub

Private Sub DefinirMaterialConcreto(ByVal sapModel As Object, ByVal nomeMaterial As String, ByVal fckMPa As Double)
    Dim ret As Long
    Dim elasticidadeTfM2 As Double

    ' Estimativa didática: Eci = 5600 * raiz(fck) em MPa, convertida para tf/m².
    elasticidadeTfM2 = 5600# * Sqr(fckMPa) * MPA_TO_TF_M2

    ret = sapModel.PropMaterial.SetMaterial(nomeMaterial, SAP_MATERIAL_CONCRETE)
    VerificarRetorno ret, "PropMaterial.SetMaterial " & nomeMaterial

    ret = sapModel.PropMaterial.SetMPIsotropic(nomeMaterial, elasticidadeTfM2, CONCRETE_POISSON, CONCRETE_THERMAL_COEFF)
    VerificarRetorno ret, "PropMaterial.SetMPIsotropic " & nomeMaterial
End Sub

Private Sub DefinirSecoes(ByVal sapModel As Object, ByRef dados As DadosPortico)
    Dim ret As Long

    RegistrarLog "INFO", "Definindo seções retangulares."

    ret = sapModel.PropFrame.SetRectangle(SECAO_PILAR, MATERIAL_CONCRETO_PILAR, dados.PilarDimYM, dados.PilarDimXM)
    VerificarRetorno ret, "PropFrame.SetRectangle " & SECAO_PILAR

    ret = sapModel.PropFrame.SetRectangle(SECAO_VIGA, MATERIAL_CONCRETO_VIGA, dados.AlturaVigaM, dados.LarguraVigaM)
    VerificarRetorno ret, "PropFrame.SetRectangle " & SECAO_VIGA
End Sub

Private Sub DefinirPadroesCarga(ByVal sapModel As Object)
    Dim ret As Long

    RegistrarLog "INFO", "Criando padrões de carga PP, SOBRECARGA e VENTO."

    ret = sapModel.LoadPatterns.Add(LOAD_PP, SAP_LOADTYPE_DEAD, 1#, True)
    VerificarRetorno ret, "LoadPatterns.Add " & LOAD_PP

    ret = sapModel.LoadPatterns.Add(LOAD_SOBRECARGA, SAP_LOADTYPE_LIVE, 0#, True)
    VerificarRetorno ret, "LoadPatterns.Add " & LOAD_SOBRECARGA

    ret = sapModel.LoadPatterns.Add(LOAD_VENTO, SAP_LOADTYPE_WIND, 0#, True)
    VerificarRetorno ret, "LoadPatterns.Add " & LOAD_VENTO
End Sub

Private Function AddFrameByCoord(ByVal sapModel As Object, ByVal x1 As Double, ByVal y1 As Double, ByVal z1 As Double, ByVal x2 As Double, ByVal y2 As Double, ByVal z2 As Double, ByVal userName As String, ByVal propName As String) As String
    Dim ret As Long
    Dim frameName As String

    frameName = vbNullString
    ret = sapModel.FrameObj.AddByCoord(x1, y1, z1, x2, y2, z2, frameName, propName, userName, SAP_COORD_GLOBAL)
    VerificarRetorno ret, "FrameObj.AddByCoord " & userName

    AddFrameByCoord = frameName
End Function

Private Sub ObterPontosFrame(ByVal sapModel As Object, ByVal frameName As String, ByRef pointI As String, ByRef pointJ As String)
    Dim ret As Long

    ret = sapModel.FrameObj.GetPoints(frameName, pointI, pointJ)
    VerificarRetorno ret, "FrameObj.GetPoints " & frameName
End Sub

Private Sub AplicarEngaste(ByVal sapModel As Object, ByVal pointName As String)
    Dim ret As Long
    Dim restr(0 To 5) As Boolean
    Dim i As Long

    For i = 0 To 5
        restr(i) = True
    Next i

    ret = sapModel.PointObj.SetRestraint(pointName, restr)
    VerificarRetorno ret, "PointObj.SetRestraint " & pointName
End Sub

Private Sub AplicarCargaDistribuidaViga(ByVal sapModel As Object, ByVal frameName As String, ByVal cargaTfM As Double)
    Dim ret As Long

    RegistrarLog "INFO", "Aplicando carga distribuída vertical na viga: " & Format$(cargaTfM, "0.00") & " tf/m."

    ret = sapModel.FrameObj.SetLoadDistributed(frameName, LOAD_SOBRECARGA, SAP_DIST_LOAD_FORCE_PER_LENGTH, SAP_DIR_GLOBAL_Z, 0#, 1#, -cargaTfM, -cargaTfM, SAP_COORD_GLOBAL, True, True)
    VerificarRetorno ret, "FrameObj.SetLoadDistributed " & frameName
End Sub

Private Sub AplicarCargaVento(ByVal sapModel As Object, ByVal pointName As String, ByVal cargaTf As Double)
    Dim ret As Long
    Dim values(0 To 5) As Double

    RegistrarLog "INFO", "Aplicando carga horizontal de vento no nó superior direito: " & Format$(cargaTf, "0.00") & " tf."

    values(0) = cargaTf
    ret = sapModel.PointObj.SetLoadForce(pointName, LOAD_VENTO, values, True, SAP_COORD_GLOBAL)
    VerificarRetorno ret, "PointObj.SetLoadForce " & pointName
End Sub

Private Sub VerificarRetorno(ByVal ret As Long, ByVal operacao As String)
    If ret <> 0 Then
        Err.Raise vbObjectError + 513, "SAP2000 OAPI", operacao & " retornou código " & CStr(ret) & "."
    End If
End Sub
