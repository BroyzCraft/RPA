Attribute VB_Name = "previa"
Sub gerar()
    
    Dim sapAberto As Variant
    sapAberto = Shell("taskkill /IM saplogon.exe", vbNormalFocus)
    
    Sheets("previa").Select
    Dim datainicio
    Dim dataFinal
    Dim usuario
    Dim senha
    datainicio = Format(form_previa_dados_gerar.datainicio, "yyyymmdd")
    dataFinal = Format(form_previa_dados_gerar.datafim, "yyyymmdd")
    usuario = form_previa_dados_gerar.usuario
    senha = form_previa_dados_gerar.senha
    
    Dim SapGui
    Dim Applic
    Dim connection
    Dim session
    Dim WSHShell
    
    'Abre o Sap instalado na sua m?quina
    Shell "C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe", vbNormalFocus
    'Inicia a vari?vel com o objeto SAP
    Set WSHShell = CreateObject("WScript.Shell")
    Do Until WSHShell.AppActivate("SAP Logon ")
        Application.Wait Now + TimeValue("0:00:01")
    Loop
    Set WSHShell = Nothing
    Set SapGui = GetObject("SAPGUI")
    Set Applic = SapGui.GetScriptingEngine
    Set connection = Applic.OpenConnection("14 - ECC PRD - EP1", True)
    Set session = connection.Children(0)
    session.findById("wnd[0]").maximize
    'DADOS PARA FAZER O LOGIN NO SISTEMA
    session.findById("wnd[0]/usr/txtRSYST-MANDT").Text = "500" 'client do sistema
    session.findById("wnd[0]/usr/txtRSYST-BNAME").Text = usuario 'usuario
    session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = senha 'senha
    session.findById("wnd[0]/usr/txtRSYST-LANGU").Text = "PT"  'idioma do sistema
    session.findById("wnd[0]").sendVKey 0 'bot?o enter para entrar no sistema
    session.findById("wnd[0]").maximize
    session.findById("wnd[0]/tbar[0]/okcd").Text = "VL10A"
    session.findById("wnd[0]").sendVKey 0
    'fechar os prints
    'sp
    session.findById("wnd[0]/tbar[1]/btn[17]").press
    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "1"
    session.findById("wnd[1]/tbar[0]/btn[2]").press
    session.findById("wnd[0]/usr/ctxtST_LEDAT-LOW").SetFocus
    session.findById("wnd[0]/usr/ctxtST_LEDAT-LOW").caretPosition = 7
    session.findById("wnd[0]").sendVKey 4
    session.findById("wnd[1]/usr/cntlCONTAINER/shellcont/shell").selectionInterval = datainicio
    session.findById("wnd[0]/usr/ctxtST_LEDAT-HIGH").SetFocus
    session.findById("wnd[0]/usr/ctxtST_LEDAT-HIGH").caretPosition = 2
    session.findById("wnd[0]").sendVKey 4
    session.findById("wnd[1]/usr/cntlCONTAINER/shellcont/shell").selectionInterval = dataFinal
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell -1, ""
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").SelectAll
    session.findById("wnd[0]/tbar[1]/btn[19]").press
    session.findById("wnd[0]/tbar[0]/btn[3]").press
    '1027
    session.findById("wnd[0]/tbar[1]/btn[17]").press
    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "2"
    session.findById("wnd[1]/tbar[0]/btn[2]").press
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell -1, ""
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").SelectAll
    session.findById("wnd[0]/tbar[1]/btn[19]").press
    session.findById("wnd[0]/tbar[0]/btn[3]").press
    'interior
    session.findById("wnd[0]/tbar[1]/btn[17]").press
    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "3"
    session.findById("wnd[1]/tbar[0]/btn[2]").press
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell -1, ""
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").SelectAll
    session.findById("wnd[0]/tbar[1]/btn[19]").press
    session.findById("wnd[0]/tbar[0]/btn[3]").press
    'rj
    session.findById("wnd[0]/tbar[1]/btn[17]").press
    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "5"
    session.findById("wnd[1]/tbar[0]/btn[2]").press
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell -1, ""
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").SelectAll
    session.findById("wnd[0]/tbar[1]/btn[19]").press
    session.findById("wnd[0]/tbar[0]/btn[3]").press
    
    Set session = Nothing
    Application.Wait Now + TimeValue("0:00:05")
    connection.CloseSession ("ses[0]")
    Set connection = Nothing
    Set sap = Nothing

End Sub

Sub atualiza()

    Sheets("previa").Visible = True
    Sheets("previa-sp").Visible = True
    Sheets("previa-retira").Visible = True
    Sheets("previa-loja").Visible = True
    Sheets("previa-rj").Visible = True
    
    Dim sapAberto As Variant
    sapAberto = Shell("taskkill /IM saplogon.exe", vbNormalFocus)
    
    Sheets("previa").Select
    Dim datainicio
    Dim dataFinal
    Dim usuario
    Dim senha
    datainicio = Format(form_previa_dados_atualizar.datainicio, "yyyymmdd")
    dataFinal = Format(form_previa_dados_atualizar.datafim, "yyyymmdd")
    usuario = form_previa_dados_atualizar.usuario
    senha = form_previa_dados_atualizar.senha
    
    Dim SapGui
    Dim Applic
    Dim connection
    Dim session
    Dim WSHShell
    
    'Abre o Sap instalado na sua m?quina
    Shell "C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe", vbNormalFocus
    'Inicia a vari?vel com o objeto SAP
    Set WSHShell = CreateObject("WScript.Shell")
    Do Until WSHShell.AppActivate("SAP Logon ")
        Application.Wait Now + TimeValue("0:00:01")
    Loop
    Set WSHShell = Nothing
    Set SapGui = GetObject("SAPGUI")
    Set Applic = SapGui.GetScriptingEngine
    Set connection = Applic.OpenConnection("14 - ECC PRD - EP1", True)
    Set session = connection.Children(0)
    session.findById("wnd[0]").maximize
    'DADOS PARA FAZER O LOGIN NO SISTEMA
    session.findById("wnd[0]/usr/txtRSYST-MANDT").Text = "500" 'client do sistema
    session.findById("wnd[0]/usr/txtRSYST-BNAME").Text = usuario 'usuario
    session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = senha 'senha
    session.findById("wnd[0]/usr/txtRSYST-LANGU").Text = "PT"  'idioma do sistema
    session.findById("wnd[0]").sendVKey 0 'bot?o enter para entrar no sistema
    session.findById("wnd[0]").maximize
    session.findById("wnd[0]/tbar[0]/okcd").Text = "VL10A"
    session.findById("wnd[0]").sendVKey 0
    'Coleta as informa??es de pedido e peso
    'SP
    session.findById("wnd[0]/tbar[1]/btn[25]").press
    session.findById("wnd[0]/usr/txtERNAM-LOW").Text = ""
    session.findById("wnd[0]/usr/txtERNAM-LOW").SetFocus
    session.findById("wnd[0]/usr/txtERNAM-LOW").caretPosition = 0
    session.findById("wnd[0]/usr/btn%_VSTEL_%_APP_%-VALU_PUSH").press
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "100F"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").Text = "100G"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").Text = "100I"
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    session.findById("wnd[0]/usr/ctxtERDAT-LOW").SetFocus
    session.findById("wnd[0]/usr/ctxtERDAT-LOW").caretPosition = 0
    session.findById("wnd[0]").sendVKey 4
    session.findById("wnd[1]/usr/cntlCONTAINER/shellcont/shell").selectionInterval = datainicio
    session.findById("wnd[0]/usr/ctxtERDAT-HIGH").SetFocus
    session.findById("wnd[0]/usr/ctxtERDAT-HIGH").caretPosition = 0
    session.findById("wnd[0]").sendVKey 4
    session.findById("wnd[1]/usr/cntlCONTAINER/shellcont/shell").selectionInterval = dataFinal
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").Select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    'session.findById("wnd[1]/usr/ctxtDY_PATH").Text = "C:\SAP"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "sp.txt"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 6
    session.findById("wnd[1]/tbar[0]/btn[11]").press
    session.findById("wnd[0]/tbar[0]/btn[3]").press
    'RETIRA
    session.findById("wnd[0]/usr/btn%_VSTEL_%_APP_%-VALU_PUSH").press
    session.findById("wnd[1]/tbar[0]/btn[16]").press
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "100b"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").Text = "100c"
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").Select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "retira.txt"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 6
    session.findById("wnd[1]/tbar[0]/btn[11]").press
    session.findById("wnd[0]/tbar[0]/btn[3]").press
    'LOJA
    session.findById("wnd[0]/usr/btn%_VSTEL_%_APP_%-VALU_PUSH").press
    session.findById("wnd[1]/tbar[0]/btn[16]").press
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "100h"
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").Select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "loja.txt"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 6
    session.findById("wnd[1]/tbar[0]/btn[11]").press
    session.findById("wnd[0]/tbar[0]/btn[3]").press
    'RJ
    session.findById("wnd[0]/usr/btn%_VSTEL_%_APP_%-VALU_PUSH").press
    session.findById("wnd[1]/tbar[0]/btn[16]").press
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "100d"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").Text = "100e"
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").Select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "rj.txt"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 6
    session.findById("wnd[1]/tbar[0]/btn[11]").press
    session.findById("wnd[0]/tbar[0]/btn[3]").press
    session.findById("wnd[0]/tbar[0]/btn[3]").press
    session.findById("wnd[0]/tbar[0]/btn[3]").press
    
    'FECHA CONEX?O SAP
    Set session = Nothing
    Application.Wait Now + TimeValue("0:00:05")
    connection.CloseSession ("ses[0]")
    Set connection = Nothing
    Set sap = Nothing
    
    AjustaDados

End Sub

Sub AjustaDados()

    'reseta e importa os dados - sp
    Sheets("previa-sp").Select
    Cells.Select
    Selection.ClearContents
    Range("A1").Select
    
    With ActiveSheet.QueryTables.Add(connection:= _
        "TEXT;" & Environ("USERPROFILE") & "\Documents\SAP\SAP GUI\sp.txt", Destination:= _
        Range("$A$1"))
        '.CommandType = 0
        .Name = "previa_1"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 932
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = True
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    Columns("A:A").Select
    Selection.Delete shift:=xlToLeft
    Rows("1:3").Select
    Selection.Delete shift:=xlUp
    Rows("2:2").Select
    Selection.Delete shift:=xlUp
    Range("A1").Select
    
    'retira
    Sheets("previa-retira").Select
    Cells.Select
    Selection.ClearContents
    Range("A1").Select
    
    With ActiveSheet.QueryTables.Add(connection:= _
        "TEXT;" & Environ("USERPROFILE") & "\Documents\SAP\SAP GUI\retira.txt", Destination:=Range("$A$1"))
        '.CommandType = 0
        .Name = "previa_1"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 932
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = True
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    Columns("A:A").Select
    Selection.Delete shift:=xlToLeft
    Rows("1:3").Select
    Selection.Delete shift:=xlUp
    Rows("2:2").Select
    Selection.Delete shift:=xlUp
    Range("A1").Select
    
    'RJ
    Sheets("previa-rj").Select
    Cells.Select
    Selection.ClearContents
    Range("A1").Select
    
    With ActiveSheet.QueryTables.Add(connection:= _
        "TEXT;" & Environ("USERPROFILE") & "\Documents\SAP\SAP GUI\rj.txt", Destination:= _
        Range("$A$1"))
        '.CommandType = 0
        .Name = "previa_1"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 932
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = True
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    Columns("A:A").Select
    Selection.Delete shift:=xlToLeft
    Rows("1:3").Select
    Selection.Delete shift:=xlUp
    Rows("2:2").Select
    Selection.Delete shift:=xlUp
    Range("A1").Select
    
    'lj
    Sheets("previa-loja").Select
    Cells.Select
    Selection.ClearContents
    Range("A1").Select
    
    With ActiveSheet.QueryTables.Add(connection:= _
        "TEXT;" & Environ("USERPROFILE") & "\Documents\SAP\SAP GUI\loja.txt", Destination:= _
        Range("$A$1"))
        '.CommandType = 0
        .Name = "previa_1"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = 932
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = True
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With
    Columns("A:A").Select
    Selection.Delete shift:=xlToLeft
    Rows("1:3").Select
    Selection.Delete shift:=xlUp
    Rows("2:2").Select
    Selection.Delete shift:=xlUp
    Range("A1").Select
        
    'extrai os dados
    Sheets("previa").Select
    Range("O8").Select
    ActiveCell.FormulaR1C1 = "=SUM('previa-sp'!C[-11])"
    Range("P8").Select
    ActiveCell.FormulaR1C1 = "=SUM('previa-sp'!C[-9])"
    
    Range("O9").Select
    ActiveCell.FormulaR1C1 = "=SUM('previa-retira'!C[-11])"
    Range("P9").Select
    ActiveCell.FormulaR1C1 = "=SUM('previa-retira'!C[-9])"
    
    Range("O10").Select
    ActiveCell.FormulaR1C1 = "=SUM('previa-loja'!C[-11])"
    Range("P10").Select
    ActiveCell.FormulaR1C1 = "=SUM('previa-loja'!C[-9])"
    
    Range("O11").Select
    ActiveCell.FormulaR1C1 = "=SUM('previa-rj'!C[-11])"
    Range("P11").Select
    ActiveCell.FormulaR1C1 = "=SUM('previa-rj'!C[-9])"
    
End Sub
