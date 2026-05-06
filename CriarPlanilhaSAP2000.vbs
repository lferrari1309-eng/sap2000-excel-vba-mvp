Option Explicit

' Cria a pasta de trabalho SAP2000_Portico_MVP.xlsm no mesmo diretório deste script.
' Requisito do Excel: habilitar "Trust access to the VBA project object model"
' para permitir a importação automática do módulo .bas.

Const XL_OPENXML_WORKBOOK_MACRO_ENABLED = 52

Dim fso, pastaScript, caminhoBas, caminhoXlsm
Dim excelApp, wb

Set fso = CreateObject("Scripting.FileSystemObject")
pastaScript = fso.GetParentFolderName(WScript.ScriptFullName)
caminhoBas = fso.BuildPath(pastaScript, "Sap2000ExcelMvp.bas")
caminhoXlsm = fso.BuildPath(pastaScript, "SAP2000_Portico_MVP.xlsm")

If Not fso.FileExists(caminhoBas) Then
    WScript.Echo "Arquivo VBA não encontrado: " & caminhoBas
    WScript.Quit 1
End If

On Error Resume Next
Set excelApp = CreateObject("Excel.Application")
If Err.Number <> 0 Then
    WScript.Echo "Não foi possível abrir o Excel. Verifique se o Microsoft Excel está instalado." & vbCrLf & Err.Description
    WScript.Quit 1
End If
On Error GoTo 0

excelApp.Visible = True
excelApp.DisplayAlerts = False

Set wb = excelApp.Workbooks.Add

On Error Resume Next
wb.VBProject.VBComponents.Import caminhoBas
If Err.Number <> 0 Then
    WScript.Echo "Não foi possível importar o módulo VBA." & vbCrLf & _
        "No Excel, habilite: File > Options > Trust Center > Trust Center Settings > Macro Settings > Trust access to the VBA project object model." & vbCrLf & _
        "Detalhe: " & Err.Description
    wb.Close False
    excelApp.Quit
    WScript.Quit 1
End If
On Error GoTo 0

If fso.FileExists(caminhoXlsm) Then
    fso.DeleteFile caminhoXlsm, True
End If

wb.SaveAs caminhoXlsm, XL_OPENXML_WORKBOOK_MACRO_ENABLED
excelApp.Run "'" & wb.Name & "'!CriarLayout"
wb.Save
excelApp.DisplayAlerts = True

WScript.Echo "Planilha criada com sucesso:" & vbCrLf & caminhoXlsm
