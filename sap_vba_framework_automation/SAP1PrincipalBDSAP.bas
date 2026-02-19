Attribute VB_Name = "SAP1PrincipalBDSAP"
'modPrincipal
'botão e fluxo principal

---> Nesse código VBA está modularizado todos os robustos processos, cada um com sua função.<---
Sub Atualizar_Dados()

    Dim session As Object
    Dim caminhoArquivo As String

    ' 1. Conectar ao SAP
    Set session = ConectarSAP()

    If session Is Nothing Then
        MsgBox "Abra o SAP, ou  retorne ao Menu Principal do servidor.", vbCritical
        Exit Sub
    End If

    ' 2. Abrir transação / tela inicial
    AbrirTransacao session

    ' 3. Preencher datas escolhidas pelo usuário
    PreencherDatas session

    ' 4. Executar relatório e exportar
    ExportarRelatorio session


    'MsgBox "Dados atualizados com sucesso!", vbInformation

End Sub

