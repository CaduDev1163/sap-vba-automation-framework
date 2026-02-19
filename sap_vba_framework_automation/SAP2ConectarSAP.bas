Attribute VB_Name = "SAP2ConectarSAP"
'modSAP
'conexão/interação direta com o servidor SAP
Function ConectarSAP() As Object

    Dim SapGuiAuto As Object
    Dim SAPApp As Object
    Dim SAPCon As Object

    On Error Resume Next

    Set SapGuiAuto = GetObject("SAPGUI")
    Set SAPApp = SapGuiAuto.GetScriptingEngine
    Set SAPCon = SAPApp.Children(0)
    Set ConectarSAP = SAPCon.Children(0)

    On Error GoTo 0

End Function

Sub AbrirTransacao(session As Object)

' Tenta voltar para raiz
    session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
    session.findById("wnd[0]").sendVKey 0
    
    Application.Wait Now + TimeValue("00:00:01")
    
    ' Se existir botão de executar Easy Access, clique
    If session.findById("wnd[0]/usr", False) Is Nothing = False Then
        session.findById("wnd[0]/usr/btnSTARTBUTTON").press
    End If
    
    On Error GoTo 0


	Os scripts .VBS do SAP são variantes, ou seja, podem mudar de acordo com a intenção do projeto.

    'Abre o caminho especifico no SAP
    ---> Script .VBS do SAP que é usado como mapa pelo VBA para abrir o caminho específico no SAP <---
    
    'Seleciona a variante e exibe as ações/agendamento
    ---> Script .VBS do SAP que é usado como mapa pelo VBA para preencher campos de textos específicos no SAP <---


End Sub


