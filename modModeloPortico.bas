Attribute VB_Name = "modModeloPortico"
Option Explicit

' Tipos e constantes do pórtico MVP.

Public Const ENTRADA_SHEET_NAME As String = "Entrada"

Public Const CM_TO_M As Double = 0.01
Public Const MPA_TO_TF_M2 As Double = 101.9716213

Public Const MATERIAL_CONCRETO_VIGA As String = "CONC_VIGA"
Public Const MATERIAL_CONCRETO_PILAR As String = "CONC_PILAR"
Public Const SECAO_VIGA As String = "VIGA_RET"
Public Const SECAO_PILAR As String = "PILAR_RET"

Public Const LOAD_PP As String = "PP"
Public Const LOAD_SOBRECARGA As String = "SOBRECARGA"
Public Const LOAD_VENTO As String = "VENTO"

Public Const FRAME_PILAR_ESQUERDO As String = "PILAR_ESQ"
Public Const FRAME_PILAR_DIREITO As String = "PILAR_DIR"
Public Const FRAME_VIGA_SUPERIOR As String = "VIGA_SUP"

Public Type DadosPortico
    VaoCm As Double
    PeDireitoCm As Double
    LarguraVigaCm As Double
    AlturaVigaCm As Double
    FckVigaMPa As Double
    PilarDimXCm As Double
    PilarDimYCm As Double
    FckPilarMPa As Double
    CargaDistribuidaTfM As Double
    CargaVentoTf As Double

    VaoM As Double
    PeDireitoM As Double
    LarguraVigaM As Double
    AlturaVigaM As Double
    PilarDimXM As Double
    PilarDimYM As Double
End Type

Public Sub ConverterDadosParaUnidadesSAP(ByRef dados As DadosPortico)
    dados.VaoM = dados.VaoCm * CM_TO_M
    dados.PeDireitoM = dados.PeDireitoCm * CM_TO_M
    dados.LarguraVigaM = dados.LarguraVigaCm * CM_TO_M
    dados.AlturaVigaM = dados.AlturaVigaCm * CM_TO_M
    dados.PilarDimXM = dados.PilarDimXCm * CM_TO_M
    dados.PilarDimYM = dados.PilarDimYCm * CM_TO_M
End Sub
