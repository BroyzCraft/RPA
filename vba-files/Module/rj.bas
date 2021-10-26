Attribute VB_Name = "rj"
Sub apagar()
    
    ' Confirma se realmente deseja resetar a tabela
    Dim confirmacao As VbMsgBoxResult
    confirmacao = MsgBox("Deseja apagar todos os registros?", vbYesNo)
    
    If confirmacao = vbYes Then
        
        'dados da planilha de apoio
        Range("B2:B8").Select
        Selection.ClearContents
        Range("I2:J2").Select
        Selection.ClearContents
        Range("I4:J5").Select
        Selection.ClearContents
        Range("H9").Select
        Selection.ClearContents
        Range("O3:R100").Select
        Selection.ClearContents
        Range("P2:P3").Select
        Selection.ClearContents
        Range("U:Y").Select
        Selection.ClearContents
        Range("AA2:AA100").Select
        Selection.ClearContents
    
        Sheets("rj-capa-corte").Select
        Range("C14:M41").Select
        Selection.ClearContents
        
    End If
    
    Sheets("rj-menu").Select
    Range("A1").Select
    
End Sub

Sub imprimirCortes()

    ' CONFIRMAÇÃO
    Dim confirmacao As VbMsgBoxResult
    confirmacao = MsgBox("Você solicitou a impressão das capas de corte, Continuar?", vbYesNo)
    
    If confirmacao = vbYes Then
        qtd = Application.InputBox("Digite quantas capas deseja imprimir: ")
        Sheets("rj-capa-corte").Select
        Range("A1:M43").Select
        Selection.PrintOut Copies:=qtd, Collate:=True
    End If
    
    confirmacao = MsgBox("Deseja gerar os arquivos CLI E PED? (mantenha o SAP fechado)", vbYesNo)
    
    If confirmacao = vbYes Then
        SAP_CLIPED
    End If
    
    Sheets("rj-menu").Select
    
End Sub

Sub imprimirControle()

    Dim confirmacao As VbMsgBoxResult
    confirmacao = MsgBox("Você solicitou a impressão do controle, Continuar?", vbYesNo)
    
    If confirmacao = vbYes Then
    
        Dim nome As String
        
        confirmacao = MsgBox("Deseja criar um novo controle? ", vbYesNo)
        Sheets("rj-menu").Select
        nome = Range("B12").Value
        
        If confirmacao = vbYes Then
            Worksheets("rj-controle").Copy after:=Worksheets(1)
            ActiveSheet.Name = nome
        End If
        
        qtd = Application.InputBox("Digite quantas capas deseja imprimir: ")
        Sheets(nome).Select
        Range("A1:J40").Select
        Selection.PrintOut Copies:=qtd, Collate:=True
    
    Else: Exit Sub
    End If
    
    confirmacao = MsgBox("Deseja salvar os dados?", vbYesNo)
    
    If confirmacao = vbYes Then
            
        'Consolida os dados
        Sheets(nome).Select
        Range("A1:J40").Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Range("A1:J40").Select
        
        'Gerar PDF
        strPathNome = "L:\Logistica\Transporte\2_ROUTEASY\0 - ARQUIVOS DA ROTEIRIZAÇÃO (EXCEL)\" & "Resumo RJ - " & nome
        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=strPathNome, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, _
        OpenAfterPublish:=False
        
    End If
    
    gerarBackup
    Sheets("rj-controle").Select
    
End Sub

Sub SAP_CLIPED()

    Dim sapAberto As Variant
    sapAberto = Shell("taskkill /IM saplogon.exe", vbNormalFocus)
    
    Dim w As Worksheet
    Dim linhaFinal As Long
    Dim aux As Long
    Dim qtdePedidos As Long
    Set w = Sheets("rj-capa-corte")
    w.Select
    linhaFinal = Range("C14").End(xlDown).Row
    qtdePedidos = linhaFinal - 13
    
    Dim SapGui
    Dim Applic
    Dim connection
    Dim session
    Dim WSHShell
    
    'Abre o Sap instalado na sua máquina
    Shell "C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe", vbNormalFocus
    
    'inicia a variável com o objeto SAP
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
    session.findById("wnd[0]/usr/txtRSYST-BNAME").Text = "BOMARQUES" 'usuario
    session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = "Leo321654987*" 'senha
    session.findById("wnd[0]/usr/txtRSYST-LANGU").Text = "PT"  'idioma do sistema
    session.findById("wnd[0]").sendVKey 0 'botão enter para entrar no sistema
    
    'gerar cli e ped
    session.findById("wnd[0]/tbar[0]/okcd").Text = "ZSDT009"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/usr/ctxtS_WERKS-LOW").Text = "*"
    session.findById("wnd[0]/usr/ctxtS_AUART-LOW").Text = "*"
    session.findById("wnd[0]/usr/ctxtS_VKBUR-LOW").Text = "*"
    session.findById("wnd[0]/usr/ctxtS_VBELN-LOW").SetFocus
    session.findById("wnd[0]/usr/ctxtS_VBELN-LOW").caretPosition = 0
    session.findById("wnd[0]/usr/btn%_S_VBELN_%_APP_%-VALU_PUSH").press
    w.Range("C14:C41").Copy
    session.findById("wnd[1]/tbar[0]/btn[16]").press
    session.findById("wnd[1]/tbar[0]/btn[24]").press
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    session.findById("wnd[0]/usr/ctxtP_VARI").Text = "/CLI CORTE"
    session.findById("wnd[0]/usr/ctxtP_VARI").SetFocus
    session.findById("wnd[0]/usr/ctxtP_VARI").caretPosition = 10
    session.findById("wnd[0]").sendVKey 8
    session.findById("wnd[0]/tbar[1]/btn[45]").press
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "cli.txt"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 3
    session.findById("wnd[1]/tbar[0]/btn[11]").press
    session.findById("wnd[0]/tbar[0]/btn[3]").press
    session.findById("wnd[0]/usr/ctxtP_VARI").Text = "/PED CORTE"
    session.findById("wnd[0]/usr/ctxtP_VARI").SetFocus
    session.findById("wnd[0]/usr/ctxtP_VARI").caretPosition = 4
    session.findById("wnd[0]").sendVKey 8
    session.findById("wnd[0]/tbar[1]/btn[45]").press
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "ped.txt"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 3
    session.findById("wnd[1]/tbar[0]/btn[11]").press
    session.findById("wnd[0]/tbar[0]/btn[3]").press
    session.findById("wnd[0]/tbar[0]/btn[3]").press
    
    'FECHA CONEXÃO SAP
    Set session = Nothing
    Application.Wait Now + TimeValue("0:00:05")
    connection.CloseSession ("ses[0]")
    Set connection = Nothing
    Set sap = Nothing
    
    Dim x As Variant
    Dim Caminho As String
    Dim arquivo As String
    Path = "C:\Users\Bruno.marques\Desktop\ConverteSapRoadNet.exe"
    x = Shell(Path, vbNormalFocus)

End Sub

Sub SAP_PRINT()

    Dim SapGui
    Dim Applic
    Dim connection
    Dim session
    Dim WSHShell
    
    'Abre o Sap instalado na sua máquina
    Shell "C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe", vbNormalFocus
    
    'inicia a variável com o objeto SAP
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
    session.findById("wnd[0]/usr/txtRSYST-BNAME").Text = "BOMARQUES" 'usuario
    session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = "Leo321654987*" 'senha
    session.findById("wnd[0]/usr/txtRSYST-LANGU").Text = "PT"  'idioma do sistema
    session.findById("wnd[0]").sendVKey 0 'botão enter para entrar no sistema
    
    'GERAR REMESSAS DO RJ
    session.findById("wnd[0]/tbar[0]/okcd").Text = "VL10A"
    session.findById("wnd[0]").sendVKey 0
    session.findById("wnd[0]/tbar[1]/btn[17]").press
    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").setCurrentCell 5, "TEXT"
    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "5"
    session.findById("wnd[1]/tbar[0]/btn[2]").press
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell -1, ""
    session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").SelectAll
    session.findById("wnd[0]/tbar[1]/btn[19]").press
    session.findById("wnd[0]/tbar[0]/btn[3]").press
    session.findById("wnd[0]/tbar[0]/btn[3]").press
    
    'FECHA CONEXÃO SAP
    Set session = Nothing
    Application.Wait Now + TimeValue("0:00:05")
    connection.CloseSession ("ses[0]")
    Set connection = Nothing
    Set sap = Nothing
    
End Sub

Sub gerarBackup()

    Dim nome As String
    Dim plan As String
    Dim macro As String
    
    Sheets("rj-menu").Select
    nome = Range("B12").Value
    plan = "10.OUTUBRO.xlsx"
    macro = "RPAs - Bruno.xlsm"
    
    Workbooks.Open ("\\Ecfs1\leo\Logistica\Transporte\4_ROTEIRIZACAO\Roteirização TP  RJ\2021\" & plan)
    
    Workbooks(macro).Activate
    Sheets(nome).Select
    ActiveSheet.Move Before:=Workbooks(plan).Sheets(1)
    
    Workbooks(plan).Close True

End Sub
