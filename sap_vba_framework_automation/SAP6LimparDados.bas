Attribute VB_Name = "SAP6LimparDados"

---> Botão no EXCEL (macro) que servirá para excluir todos os dados colados da planilha EXPORT para a planilha SAP. <---
Sub LimparDadosTabela()

    Dim ws As Worksheet
    Dim ultimaLinha As Long
    Dim ultimaColuna As Long
    Dim rngLimpeza As Range

    Set ws = ActiveSheet 

    ' Se não houver nada em C6, sai
    If IsEmpty(ws.Range("C6")) Then
        MsgBox "Só deletará conteúdo a partir da célula C6.", vbInformation
        Exit Sub
    End If

---> Lógica para apagar conteúdo até a última linha da tabela de dados. <---

    ' Descobre última linha e coluna usadas
    ultimaLinha = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    ultimaColuna = ws.Cells(6, ws.Columns.Count).End(xlToLeft).Column

    ' Define intervalo dinâmico
    Set rngLimpeza = ws.Range(ws.Cells(3, 3), ws.Cells(ultimaLinha, ultimaColuna))

    ' Apaga TUDO (conteúdo + formatação)
    rngLimpeza.Clear

    MsgBox "Deletado com sucesso.", vbInformation


End Sub

