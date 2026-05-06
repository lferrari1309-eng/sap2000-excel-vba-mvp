Attribute VB_Name = "modPrincipal"
Option Explicit

' Módulo principal: orquestra leitura, validação, conexão com SAP2000 e geração do modelo.

Public Sub ExportarParaSAP2000()
    On Error GoTo TratarErro

    Dim dados As DadosPortico
    Dim mensagemErro As String
    Dim sapObject As Object
    Dim sapModel As Object

    EnsureLogSheet
    RegistrarLog "INFO", "Iniciando exportação para SAP2000."

    If Not LerDadosEntrada(dados) Then
        MsgBox "Não foi possível ler os dados da aba Entrada. Verifique se o layout foi criado e se os campos são numéricos.", vbExclamation, "SAP2000 MVP"
        Exit Sub
    End If

    If Not ValidarDadosEntrada(dados, mensagemErro) Then
        RegistrarLog "ERRO", "Validação falhou:" & vbCrLf & mensagemErro
        MsgBox "Corrija os campos antes de exportar:" & vbCrLf & vbCrLf & mensagemErro, vbExclamation, "SAP2000 MVP"
        Exit Sub
    End If

    RegistrarLog "INFO", "Entradas validadas. Vão = " & Format$(dados.VaoM, "0.000") & " m; pé-direito = " & Format$(dados.PeDireitoM, "0.000") & " m."

    If Not ConectarOuAbrirSAP2000(sapObject, sapModel) Then
        MsgBox "Não foi possível conectar ou abrir o SAP2000. Confira se o SAP2000 está instalado e registrado no Windows.", vbCritical, "SAP2000 MVP"
        Exit Sub
    End If

    If Not CriarModeloPorticoSAP2000(sapModel, dados) Then
        MsgBox "A conexão com o SAP2000 foi feita, mas o modelo não pôde ser criado. Consulte a aba Log.", vbCritical, "SAP2000 MVP"
        Exit Sub
    End If

    MsgBox "Pórtico criado com sucesso no SAP2000.", vbInformation, "SAP2000 MVP"
    Exit Sub

TratarErro:
    RegistrarLog "ERRO", "Erro inesperado na rotina principal: " & Err.Description
    MsgBox "Erro inesperado: " & Err.Description & vbCrLf & "Consulte a aba Log para detalhes.", vbCritical, "SAP2000 MVP"
End Sub
