Attribute VB_Name = "SAP5ColarDados"
Sub ColarDados()

    Dim wb As Workbook
    Dim wbExport As Workbook
    Dim wsDestino As Worksheet
    Dim wsOrigem As Worksheet
    Dim ultimaLinha As Long
    Dim intervaloOrigem As Range
    Dim encontrado As Boolean
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Procurar o arquivo EXPORT aberto
---> Os arquivos contendo dados baixados do SAP são salvos com nome EXPORT e logo após o download são verificados se estão abertos ou não <---

    For Each wb In Workbooks
        If wb.Name Like "EXPORT*.xlsx" Then
            Set wbExport = wb
            encontrado = True
            Exit For
        End If
    Next wb
    
    If Not encontrado Then
        MsgBox "Arquivo EXPORT não está aberto.", vbExclamation
        Exit Sub
    End If

	---> Váriaveis de tipo Workbook recebem as planilhas onde o arquivo EXPORT será colado e onde o arquivo export esta aberto. <---
    Set wsDestino = ThisWorkbook.Sheets("SAP")
    Set wsOrigem = wbExport.Sheets(1)

---> Lógica robusta para verificar até onde os dados da planilha EXPORT acaba e processo de colar esses dados para a planilha Destino. Após esse processo o VBA fechará o arquivo EXPORT <---
    
    ' Descobrir última linha preenchida na planilha SAP
    ultimaLinha = wsDestino.Cells(wsDestino.Rows.Count, 1).End(xlUp).Row
    
    ' Define intervalo de origem
    Set intervaloOrigem = wsOrigem.UsedRange
    
    ' Copia dados para abaixo da última linha
    intervaloOrigem.Copy
    wsDestino.Cells(ultimaLinha + 3, 3).PasteSpecial xlPasteValues
    
    Application.CutCopyMode = False
    
    ' Ajustar largura das colunas automaticamente
    wsDestino.Columns.AutoFit
    
    ' Fechar EXPORT
    wbExport.Close SaveChanges:=False
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    MsgBox "Dados adicionados com sucesso!", vbInformation

End Sub
