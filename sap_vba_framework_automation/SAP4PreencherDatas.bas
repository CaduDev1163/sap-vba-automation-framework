Attribute VB_Name = "SAP4PreencherDatas"
'modDatas
'configuração das datas manuais por meio da própria planilha

---> Essa configuração pode não somente ser de datas, mas de preenchimentos que são necessários ser manuais usando textos contidos em células excel <---

Sub PreencherDatas(session As Object)

    Dim dataInicio As Date
    Dim dataFim As Date

    Application.Wait Now + TimeValue("00:00:03")

	--->As células específicas são mencionadas sendo inseridas em váriaveis (uma célula para cada variável) de tipo Date (que é o caso de datas) <---    
    dataInicio = Sheets("SAP").Range("D1").Value
    dataFim = Sheets("SAP").Range("D2").Value

	---> Os IDs .VBS são mencionados com um .Text pois neles serão preenchidos as datas digitadas nas células citadas acima
    session.findById("wnd[0]/usr/ctxtSO_LOW").Text = _
        Format(dataInicio, "dd.mm.yyyy")

    session.findById("wnd[0]/usr/ctxtSO_HIGH").Text = _
        Format(dataFim, "dd.mm.yyyy")

    session.findById("wnd[0]").sendVKey 0

End Sub



