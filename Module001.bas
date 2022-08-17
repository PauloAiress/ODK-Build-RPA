Attribute VB_Name = "Module001"
Sub IE03_MB52()
'essa macro é executada através de um script python
'automação da extração do relatório de equipamentos e de estoque

Application.ScreenUpdating = False

        '############### SETAR O SAP ##################
        
        If Not IsObject(SAP) Then
           Set SapGuiAuto = GetObject("SAPGUI")
           Set SAP = SapGuiAuto.GetScriptingEngine
        End If
        If Not IsObject(Connection) Then
           Set Connection = SAP.Children(0)
        End If
        If Not IsObject(session) Then
           Set session = Connection.Children(0)
        End If
        If IsObject(WScript) Then
           WScript.ConnectObject session, "on"
           WScript.ConnectObject SAP, "on"
        End If
        
        Dim obras(7) As String
        obras(0) = "A329"
        obras(1) = "A331"
        obras(2) = "A339"
        obras(3) = "A341"
        obras(4) = "A342"
        obras(5) = "A343"
        obras(6) = "A345"
        obras(7) = "A347"
        
        'lista de verificação de qual obra será atualizada
        
        Dim validador_obras(7) As String
        validador_obras(0) = False
        validador_obras(1) = True
        validador_obras(2) = False
        validador_obras(3) = True
        validador_obras(4) = False
        validador_obras(5) = False
        validador_obras(6) = False
        validador_obras(7) = False
        
        '################################################ LOOP PARA ATUALIZAR OS FORMULÁRIOS DE OBRAS ATIVAS ################################################
        For i = 0 To 7
            If validador_obras(i) = True Then
                Set workb = ThisWorkbook
                workb_name = ActiveWorkbook.Name
                workb.Sheets("IE03").Select
                Cells.Select
                Selection.ClearContents
                obra = Right(obras(i), 3)
                data_hj = Replace(Date, "/", "-")
                
                '################### IE03 ########################
                
                session.findById("wnd[0]/tbar[0]/okcd").Text = "/NIE03"
                session.findById("wnd[0]").sendVKey 0
                session.findById("wnd[0]").sendVKey 4
                'se for a 331, buscar todas as divisões ativas com equipamentos
                If obras(i) = "A331" Then
                    session.findById("wnd[0]/usr/ctxtGSBER-LOW").Text = ""
                    session.findById("wnd[0]/usr/btn%_GSBER_%_APP_%-VALU_PUSH").press
                    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "A331"
                    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").Text = "A344"
                    session.findById("wnd[0]").sendVKey 8
                Else
                    session.findById("wnd[0]/usr/ctxtGSBER-LOW").Text = obras(i)
                End If
                session.findById("wnd[0]").sendVKey 8
                session.findById("wnd[0]/tbar[1]/btn[16]").press
                session.findById("wnd[1]/tbar[0]/btn[0]").press
                session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[0,0]").Select
                session.findById("wnd[1]/tbar[0]/btn[0]").press
                session.findById("wnd[1]/tbar[0]/btn[0]").press
                Set wb = Workbooks("Planilha em Basis (1)")
                wb.Sheets("Plan1").Range("A1:Q2000").Copy Destination:=workb.Sheets("IE03").Range("A1")
                session.findById("wnd[1]/tbar[0]/btn[0]").press
                Windows(workb_name).Activate
                ActiveWindow.WindowState = xlMaximized
                Sheets("IE03").Select
                
                'filtrar os itens que não são equipamentos (BA, CAP, CX, DIF, GD, MD, PN, TQ, TR)
                
                Range("C1").Select
                Selection.End(xlDown).Select
                num_linha = ActiveCell.Row
                Selection.End(xlUp).Select
                Range("K1").Select
                ActiveSheet.Range("K1").Value = "Filtro"
                ActiveSheet.Range("K2").Formula = "=or(C2=""BA"",C2=""CAP"",C2=""CX"",C2=""DIF"", C2=""GD"",C2=""MD"",C2=""PN"",C2=""TQ"",C2=""TR"")"
                ActiveSheet.Range("K2:K" & num_linha).Formula = ActiveSheet.Range("K2").Formula
                Range("K1").Select
                Selection.AutoFilter
                ActiveSheet.Range("$A$1:$K$" & num_linha).AutoFilter Field:=11, Criteria1:="VERDADEIRO"
                Rows("2:2").Select
                Range(Selection, Selection.End(xlDown)).Select
                Selection.SpecialCells(xlCellTypeVisible).Select
                Selection.Delete Shift:=xlUp
                Selection.AutoFilter
                Range("A1").Select
                
                '#################################### TIPOS DE "CHOICES" ############################################
                
                'tipos de "choices" XLSForm ODK
                'choices_ZPM005-Apontamento_Prefixo_do_Equipamento                                                             >> IE03
                'choices_ZPM005-Apontamento_Tipo_Combustivel                                                                            >>  45211
                'choices_ZPM005-Apontamento_apontamento_lubrificantes_compartimento                       >> Compartimentos
                'choices_ZPM005-Apontamento_apontamento_lubrificantes_material                                      >> MB52
                'choices_ZPM005-Apontamento_apontamento_lubrificantes_tipo                                               >> R, RM, T
                
                '###################################################################################################
                
                'abrir o arquivo do formulário XLSForm - ODK
                
                Workbooks.Open "G:\Meu Drive\Teste_ODK\Forms\" & obra & "\Form-Apontamento-" & obra & ".xlsx"
                Set workb01 = ActiveWorkbook
                workb01_name = ActiveWorkbook.Name
                'limpar os dados antigos da planilha
                workb01.Sheets("choices").Select
                Range("A2:Z3000").Select
                Selection.ClearContents
                Cells(1, 1).Value = "list name"
                Cells(1, 2).Value = "name"
                Cells(1, 3).Value = "label::English"
                'copiar dados do relatório SAP
                workb.Sheets("IE03").Range("B2:D2000").Copy Destination:=workb01.Sheets("choices").Range("B2")
                
                'formatar equipamentos
                
                workb01.Sheets("choices").Range("B2").Select
                Selection.End(xlDown).Select
                linha_num = ActiveCell.Row
                Selection.End(xlUp).Select
                workb01.Sheets("choices").Range("C2").Select
                Selection.NumberFormat = "General"
                ActiveCell.FormulaR1C1 = "=RC[-1]&"" = ""&RC[1]"
                ActiveSheet.Range("C2:C" & linha_num).Formula = ActiveSheet.Range("C2").Formula
                Columns("C:C").EntireColumn.AutoFit
                workb01.Sheets("choices").Range("C2:D" & linha_num).Select
                Selection.Copy
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                    :=False, Transpose:=False
                workb01.Sheets("choices").Range("D2").Select
                workb01.Sheets("choices").Range(Selection, Selection.End(xlDown)).Select
                Selection.ClearContents
                workb01.Sheets("choices").Range("B2").Select
                
                'informar a choice dos equipamentos
                
                For Each valor In workb01.Sheets("choices").Range("A2:A" & linha_num)
                    linha = valor.Row
                    Cells(linha, 1).Value = "choices_ZPM005-Apontamento-" & obra & "_Prefixo_do_Equipamento"
                Next
                
                'informar a choice do combustível
                
                linha = linha + 1
                Cells(linha, 1).Value = "choices_ZPM005-Apontamento-" & obra & "_Tipo_Combustivel"
                Cells(linha, 2).Value = 45211
                Cells(linha, 3).Value = "45211 = OLEO DIESEL"
                
                
                '########################## COMPARTIMENTOS ##############################
                
                'CARCACA_DO_AMORTECEDOR, CARTER_DO_MOTOR, COMANDO_FINAL_DIREITO
                'COMANDO_FINAL_ESQUERDO, GRAXA, HIDRAULICO, RODA_GUIA, TRANSMISSAO
                'SISTEMA DE ARREFECIMENTO, UNIDADE COMPRESSOR, REDUTOR DE GIRO
        
                '##########################################################################
                
                'inserir os compartimentos
        
                linha = linha + 1
                Dim compartimentos(10) As String
                compartimentos(0) = "CARCACA DO AMORTECEDOR"
                compartimentos(1) = "CARTER DO MOTOR"
                compartimentos(2) = "COMANDO FINAL DIREITO"
                compartimentos(3) = "COMANDO FINAL ESQUERDO"
                compartimentos(4) = "GRAXA"
                compartimentos(5) = "HIDRAULICO"
                compartimentos(6) = "RODA GUIA"
                compartimentos(7) = "TRANSMISSAO"
                compartimentos(8) = "SISTEMA DE ARREFECIMENTO"
                compartimentos(9) = "UNIDADE COMPRESSOR"
                compartimentos(10) = "REDUTOR DE GIRO"
                
                For e = 0 To 10
                    Cells(linha, 3).Value = compartimentos(e)
                    Cells(linha, 2).Value = Replace(compartimentos(e), " ", "_")
                    Cells(linha, 1).Value = "choices_ZPM005-Apontamento-" & obra & "_apontamento_lubrificantes_compartimento"
                    linha = linha + 1
                Next
                
                '################### MB52 ########################
                
                Dim arr_materiais(141) As String
                'lista de todos os materiais de estoque de lubrificação/graxa
                    arr_materiais(0) = 8881
                    arr_materiais(1) = 8882
                    arr_materiais(2) = 8883
                    arr_materiais(3) = 8885
                    arr_materiais(4) = 8886
                    arr_materiais(5) = 8888
                    arr_materiais(6) = 8891
                    arr_materiais(7) = 8892
                    arr_materiais(8) = 8893
                    arr_materiais(9) = 8894
                    arr_materiais(10) = 8895
                    arr_materiais(11) = 8896
                    arr_materiais(12) = 8897
                    arr_materiais(13) = 8898
                    arr_materiais(14) = 8899
                    arr_materiais(15) = 8900
                    arr_materiais(16) = 8901
                    arr_materiais(17) = 8902
                    arr_materiais(18) = 8903
                    arr_materiais(19) = 8904
                    arr_materiais(20) = 8905
                    arr_materiais(21) = 8906
                    arr_materiais(22) = 8907
                    arr_materiais(23) = 8908
                    arr_materiais(24) = 8909
                    arr_materiais(25) = 8910
                    arr_materiais(26) = 8911
                    arr_materiais(27) = 8912
                    arr_materiais(28) = 8913
                    arr_materiais(29) = 8914
                    arr_materiais(30) = 8915
                    arr_materiais(31) = 8916
                    arr_materiais(32) = 8917
                    arr_materiais(33) = 8918
                    arr_materiais(34) = 8919
                    arr_materiais(35) = 8920
                    arr_materiais(36) = 8921
                    arr_materiais(37) = 8922
                    arr_materiais(38) = 8923
                    arr_materiais(39) = 8925
                    arr_materiais(40) = 8926
                    arr_materiais(41) = 8936
                    arr_materiais(42) = 8937
                    arr_materiais(43) = 8939
                    arr_materiais(44) = 10869
                    arr_materiais(45) = 14629
                    arr_materiais(46) = 20360
                    arr_materiais(47) = 21183
                    arr_materiais(48) = 21185
                    arr_materiais(49) = 21186
                    arr_materiais(50) = 21187
                    arr_materiais(51) = 21189
                    arr_materiais(52) = 21704
                    arr_materiais(53) = 21789
                    arr_materiais(54) = 22877
                    arr_materiais(55) = 23026
                    arr_materiais(56) = 23084
                    arr_materiais(57) = 28470
                    arr_materiais(58) = 28871
                    arr_materiais(59) = 29077
                    arr_materiais(60) = 29661
                    arr_materiais(61) = 30277
                    arr_materiais(62) = 30308
                    arr_materiais(63) = 33893
                    arr_materiais(64) = 34326
                    arr_materiais(65) = 35659
                    arr_materiais(66) = 37077
                    arr_materiais(67) = 38248
                    arr_materiais(68) = 39841
                    arr_materiais(69) = 40098
                    arr_materiais(70) = 40569
                    arr_materiais(71) = 41422
                    arr_materiais(72) = 42021
                    arr_materiais(73) = 42023
                    arr_materiais(74) = 42072
                    arr_materiais(75) = 42193
                    arr_materiais(76) = 42194
                    arr_materiais(77) = 42263
                    arr_materiais(78) = 42402
                    arr_materiais(79) = 43601
                    arr_materiais(80) = 43602
                    arr_materiais(81) = 44296
                    arr_materiais(82) = 44297
                    arr_materiais(83) = 45236
                    arr_materiais(84) = 45575
                    arr_materiais(85) = 45576
                    arr_materiais(86) = 47988
                    arr_materiais(87) = 50038
                    arr_materiais(88) = 50039
                    arr_materiais(89) = 52975
                    arr_materiais(90) = 52976
                    arr_materiais(91) = 53470
                    arr_materiais(92) = 53949
                    arr_materiais(93) = 55043
                    arr_materiais(94) = 57262
                    arr_materiais(95) = 58055
                    arr_materiais(96) = 60686
                    arr_materiais(97) = 71209
                    arr_materiais(98) = 75359
                    arr_materiais(99) = 76964
                    arr_materiais(100) = 77116
                    arr_materiais(101) = 78117
                    arr_materiais(102) = 78215
                    arr_materiais(103) = 78410
                    arr_materiais(104) = 79364
                    arr_materiais(105) = 80296
                    arr_materiais(106) = 80299
                    arr_materiais(107) = 80681
                    arr_materiais(108) = 81161
                    arr_materiais(109) = 81483
                    arr_materiais(110) = 81786
                    arr_materiais(111) = 82070
                    arr_materiais(112) = 82089
                    arr_materiais(113) = 82414
                    arr_materiais(114) = 82556
                    arr_materiais(115) = 83108
                    arr_materiais(116) = 83817
                    arr_materiais(117) = 83818
                    arr_materiais(118) = 83822
                    arr_materiais(119) = 83823
                    arr_materiais(120) = 83824
                    arr_materiais(121) = 83855
                    arr_materiais(122) = 83972
                    arr_materiais(123) = 83973
                    arr_materiais(124) = 83974
                    arr_materiais(125) = 83975
                    arr_materiais(126) = 84256
                    arr_materiais(127) = 84381
                    arr_materiais(128) = 84727
                    arr_materiais(129) = 84728
                    arr_materiais(130) = 84781
                    arr_materiais(131) = 84783
                    arr_materiais(132) = 85640
                    arr_materiais(133) = 85641
                    arr_materiais(134) = 85642
                    arr_materiais(135) = 85654
                    arr_materiais(136) = 85660
                    arr_materiais(137) = 85793
                    arr_materiais(138) = 86023
                    arr_materiais(139) = 86024
                    arr_materiais(140) = 86028
                    arr_materiais(141) = 86417
        
                 'Dim materiais As New MSForms.DataObject
                ' materiais.SetText Join(arr_materiais, vbNewLine)
                ' materiais.PutInClipboard
                 
                session.findById("wnd[0]/tbar[0]/okcd").Text = "/NMB52"
                session.findById("wnd[0]").sendVKey 0
                'se for A331, buscar todas as divisões com estoque
                If obras(i) = "A331" Then
                    session.findById("wnd[0]/usr/btn%_WERKS_%_APP_%-VALU_PUSH").press
                    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text = "A331"
                    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").Text = "A344"
                    session.findById("wnd[1]/tbar[0]/btn[8]").press
                Else
                    session.findById("wnd[0]/usr/ctxtWERKS-LOW").Text = obras(i)
                End If
                session.findById("wnd[0]/usr/ctxtMATNR-LOW").Text = ""
                session.findById("wnd[0]/usr/btn%_MATNR_%_APP_%-VALU_PUSH").press
                count_inp = 0
                count_sb = 7
                For Each valor In arr_materiais
                    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1," & count_inp & "]").Text = valor
                    count_inp = count_inp + 1
                    If count_inp = 8 Then
                        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE").verticalScrollbar.Position = count_sb
                        count_sb = count_sb + 7
                        count_inp = 1
                    End If
                Next
                session.findById("wnd[1]/tbar[0]/btn[8]").press
                session.findById("wnd[0]/tbar[1]/btn[8]").press
                session.findById("wnd[0]").sendVKey 43
                session.findById("wnd[1]/usr/ctxtDY_PATH").Text = "G:\Meu Drive\Teste_ODK\exports_SAP\" & obra
                caminho_mb52 = session.findById("wnd[1]/usr/ctxtDY_PATH").Text
                session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "MB52_" & obra & "_" & data_hj & ".XLSX"
                nome_mb52 = session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text
                session.findById("wnd[1]/tbar[0]/btn[0]").press
                'puxar o relatório em função da lista de insumos a ser enviada pelo CIL
                On Error Resume Next
                Workbooks.Open "G:\Meu Drive\Teste_ODK\exports_SAP\" & obra & "\MB52_" & obra & "_" & data_hj & ".XLSX"
                    For Each wb In Application.Workbooks
                        If ActiveWorkbook.Name = nome_mb52 Then
                            Windows(ActiveWorkbook).Activate
                            Set workb02 = ActiveWorkbook
                            Exit For
                        End If
                    Next
                workb02.Sheets("Sheet1").Activate
                Cells.Select
                Selection.ClearOutline
                Columns("A:A").Select
                ActiveSheet.Range("$A$1:$H$2000").RemoveDuplicates Columns:=1, Header:=xlYes
                workb02.Sheets("Sheet1").Range("A2:B100").Copy Destination:=workb01.Sheets("choices").Range("B" & linha)
                For Each valor In workb01.Sheets("choices").Range("B" & linha & ":B1000")
                    If valor = "" Then
                        Exit For
                    Else
                        linha_final = valor.Row
                    End If
                Next
                For Each valor In workb01.Sheets("choices").Range("A" & linha & ":A" & linha_final)
                    celula = valor.Address
                    workb01.Sheets("choices").Range(celula).Value = "choices_ZPM005-Apontamento-" & obra & "_apontamento_lubrificantes_material"
                Next
        
                 '####################### TIPOS DE APONTAMENTOS ##########################
                
                'REPOSIÇÃO, REPARO MECÂNICO, TROCA
        
                '##########################################################################
                
                'inserir os tipos de reposição
        
                linha = linha_final + 1
                Dim tipo(2) As String
                tipo(0) = "Reposição"
                tipo(1) = "Reparo Mecânico"
                tipo(2) = "Troca"
                Dim abr_tipo(2) As String
                abr_tipo(0) = "R -> "
                abr_tipo(1) = "RM -> "
                abr_tipo(2) = "T -> "
                Dim abr(2) As String
                abr(0) = "R"
                abr(1) = "RM"
                abr(2) = "T"
                
                For x = 0 To 2
                    workb01.Sheets("choices").Cells(linha, 3).Value = abr_tipo(x) & tipo(x)
                    workb01.Sheets("choices").Cells(linha, 2).Value = abr(x)
                    workb01.Sheets("choices").Cells(linha, 1).Value = "choices_ZPM005-Apontamento-" & obra & "_apontamento_lubrificantes_tipo"
                    linha = linha + 1
                Next
                
                'salvar o formulário
                workb01.Close savechanges:=True
                workb02.Close savechanges:=True
            End If
        Next
        
        Dim retval
        retval = Shell("python3 ""G:\Meu Drive\Teste_ODK\vba_py\app_convert_xml.py""", 1)
        
        workb.Close savechanges:=False
End Sub
