Attribute VB_Name = "modValidacao"
Option Explicit

' Módulo de validação das entradas do usuário.

Public Function ValidarDadosEntrada(ByRef dados As DadosPortico, ByRef mensagemErro As String) As Boolean
    mensagemErro = vbNullString

    ValidarMaiorQueZero dados.VaoCm, "Vão/comprimento do pórtico [cm]", mensagemErro
    ValidarMaiorQueZero dados.PeDireitoCm, "Pé-direito [cm]", mensagemErro
    ValidarMaiorQueZero dados.LarguraVigaCm, "Largura da viga [cm]", mensagemErro
    ValidarMaiorQueZero dados.AlturaVigaCm, "Altura da viga [cm]", mensagemErro
    ValidarMaiorQueZero dados.FckVigaMPa, "Fck da viga [MPa]", mensagemErro
    ValidarMaiorQueZero dados.PilarDimXCm, "Dimensão X do pilar [cm]", mensagemErro
    ValidarMaiorQueZero dados.PilarDimYCm, "Dimensão Y do pilar [cm]", mensagemErro
    ValidarMaiorQueZero dados.FckPilarMPa, "Fck dos pilares [MPa]", mensagemErro
    ValidarMaiorQueZero dados.CargaDistribuidaTfM, "Carga distribuída vertical na viga [tf/m]", mensagemErro
    ValidarMaiorQueZero dados.CargaVentoTf, "Carga horizontal de vento [tf]", mensagemErro

    If Len(mensagemErro) = 0 Then
        ValidarDadosEntrada = True
    Else
        ValidarDadosEntrada = False
    End If
End Function

Private Sub ValidarMaiorQueZero(ByVal valor As Double, ByVal nomeCampo As String, ByRef mensagemErro As String)
    If valor <= 0 Then
        mensagemErro = mensagemErro & "- " & nomeCampo & " deve ser maior que zero." & vbCrLf
    End If
End Sub
