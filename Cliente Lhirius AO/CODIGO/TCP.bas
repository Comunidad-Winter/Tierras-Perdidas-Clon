Attribute VB_Name = "Mod_TCP"
Option Explicit

Public UserMaxAGU As Integer
Public UserMinAGU As Integer
Public UserMaxHAM As Integer
Public UserMinHAM As Integer


Public Warping As Boolean
Public LlegaronSkills As Boolean
Public LlegaronAtrib As Boolean
Public LlegoFama As Boolean



Public Function PuedoQuitarFoco() As Boolean
PuedoQuitarFoco = True
'PuedoQuitarFoco = Not frmEstadisticas.Visible And _
'                 Not frmGuildAdm.Visible And _
'                 Not frmGuildDetails.Visible And _
'                 Not frmGuildBrief.Visible And _
'                 Not frmGuildFoundation.Visible And _
'                 Not frmGuildLeader.Visible And _
'                 Not frmCharInfo.Visible And _
'                 Not frmGuildNews.Visible And _
'                 Not frmGuildSol.Visible And _
'                 Not frmCommet.Visible And _
'                 Not frmPeaceProp.Visible
'
End Function

Sub HandleData(ByVal Rdata As String)
    On Error Resume Next
    
    Dim UserIndex As Integer
    Dim RetVal As Variant
    Dim X As Integer
    Dim Y As Integer
    Dim charindex As Integer
    Dim tempint As Integer
    Dim tempstr As String
    Dim slot As Integer
    Dim MapNumber As String
    Dim i As Integer, k As Integer
    Dim cad$, Index As Integer, m As Integer
    Dim T() As String
    
    Dim tstr As String
    Dim tstr2 As String
    
    
    Dim sData As String
    'desencriptar
    Rdata = ModDesEncrypt.DesEncriptar(Rdata)
    'desencriptar
    sData = UCase$(Rdata)
If Left$(sData, 4) = "INVI" Then CartelInvisibilidad = Right$(sData, Len(sData) - 4)
If Left$(sData, 4) = "INMO" Then CartelParalisis = Right$(sData, Len(sData) - 4)
    
    Select Case sData
        Case "LOGGED"            ' >>>>> LOGIN :: LOGGED
                   AlphaY = 130
            mode = True
            logged = True
            UserCiego = False
            'EngineRun = True
            IScombate = False
            UserDescansar = False
            Nombres = True
            Call frmMain.CargarAmigos
            If frmCrearPersonaje.Visible = True Or frmCuent.Visible = True Then
                Unload frmCrearPersonaje
                Unload frmConnect
                Unload frmCuent
                frmMain.Show
            End If
            
             Audio.StopWave
            Call SetConnected
            'Mostramos el Tip
            If tipf = "1" And PrimeraVez Then
                 Call CargarTip
                 frmtip.Visible = True
                 PrimeraVez = False
            End If
            bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 2 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 4, True, False)
            Exit Sub
        Case "QTDL"              ' >>>>> Quitar Dialogos :: QTDL
            Call Dialogos.RemoveAllDialogs
            Exit Sub
        Case "NAVEG"
            UserNavegando = Not UserNavegando
            Exit Sub
        Case "FINOK" ' Graceful exit ;))
#If UsarWrench = 1 Then
            frmMain.Socket1.Disconnect
            frmMain.Socket1.Cleanup
#Else
            If frmMain.Winsock1.State <> sckClosed Then _
                frmMain.Winsock1.Close
#End If
            frmMain.Visible = False
            logged = False
            UserParalizado = False
            IScombate = False
            pausa = False
            UserMeditar = False
            UserDescansar = False
            UserNavegando = False
            frmConnect.Visible = True
            Call Audio.StopWave
            'frmMain.IsPlaying = PlayLoop.plNone
            bRain = False
            bFogata = False
            SkillPoints = 0
    '        frmMain.Label1.Visible = False
            Call Dialogos.RemoveAllDialogs
            For i = 1 To LastChar
                charlist(i).invisible = False
            Next i
 
            bK = 0
            UserName = frmCuent.Nombre(Index)
            UserPassword = passcuent
            frmMain.Socket1.HostName = CurServerIp
            frmMain.Socket1.RemotePort = CurServerPort
            frmMain.Socket1.Connect
            Exit Sub
        '[KEVIN]-----------------------------------------------
        '**************************************************************
        Case "INITBANKO"
            frmBanco.Show , frmMain
            Exit Sub
                    Case "FINCOMOK"          ' >>>>> Finaliza Comerciar :: FINCOMOK
            frmComerciar.List1(0).Clear
            frmComerciar.List1(1).Clear
            NPCInvDim = 0
            Unload frmComerciar
            Comerciando = False
            Exit Sub
        '[KEVIN]**************************************************************
        '-----------------------------------------------------------------------------
        Case "FINBANOK"          ' >>>>> Finaliza Banco :: FINBANOK
            frmBancoObj.List1(0).Clear
            frmBancoObj.List1(1).Clear
            NPCInvDim = 0
            Unload frmBancoObj
            Comerciando = False
            Exit Sub
        '[/KEVIN]***********************************************************************
        '------------------------------------------------------------------------------
        Case "INITCOM"           ' >>>>> Inicia Comerciar :: INITCOM
            i = 1
            Do While i <= MAX_INVENTORY_SLOTS
                If Inventario.OBJIndex(i) <> 0 Then
                        frmComerciar.List1(1).AddItem Inventario.ItemName(i)
                Else
                        frmComerciar.List1(1).AddItem "Nada"
                End If
                i = i + 1
            Loop
            Comerciando = True
            frmComerciar.Show , frmMain
            Exit Sub
        '[KEVIN]-----------------------------------------------
        '**************************************************************
            
        Case "INITBANCO"           ' >>>>> Inicia Comerciar :: INITBANCO
            Dim ii As Integer
            ii = 1
            Do While ii <= MAX_INVENTORY_SLOTS
                If Inventario.OBJIndex(ii) <> 0 Then
                        frmBancoObj.List1(1).AddItem Inventario.ItemName(ii)
                Else
                        frmBancoObj.List1(1).AddItem "Nada"
                End If
                ii = ii + 1
            Loop
            
            
            i = 1
            Do While i <= UBound(UserBancoInventory)
                If UserBancoInventory(i).OBJIndex <> 0 Then
                        frmBancoObj.List1(0).AddItem UserBancoInventory(i).Name
                Else
                        frmBancoObj.List1(0).AddItem "Nada"
                End If
                i = i + 1
            Loop
            Comerciando = True
            frmBancoObj.Show , frmMain
            Exit Sub
        '---------------------------------------------------------------
        '[/KEVIN]******************
        '[Alejo]
        Case "INITCOMUSU"
            If frmComerciarUsu.List1.ListCount > 0 Then frmComerciarUsu.List1.Clear
            If frmComerciarUsu.List2.ListCount > 0 Then frmComerciarUsu.List2.Clear
            
            For i = 1 To MAX_INVENTORY_SLOTS
                If Inventario.OBJIndex(i) <> 0 Then
                        frmComerciarUsu.List1.AddItem Inventario.ItemName(i)
                        frmComerciarUsu.List1.ItemData(frmComerciarUsu.List1.NewIndex) = Inventario.Amount(i)
                Else
                        frmComerciarUsu.List1.AddItem "Nada"
                        frmComerciarUsu.List1.ItemData(frmComerciarUsu.List1.NewIndex) = 0
                End If
            Next i
            Comerciando = True
            frmComerciarUsu.Show , frmMain
        Case "FINCOMUSUOK"
            frmComerciarUsu.List1.Clear
            frmComerciarUsu.List2.Clear
            
            Unload frmComerciarUsu
            Comerciando = False
            '[/Alejo]

            Case "ERRNOM"
            
        frmMensaje.Show
            frmMensaje.MSG.Caption = "El Personaje No existe!!!"
            Exit Sub
        Case "ERRPAS"
            
        frmMensaje.Show
            frmMensaje.MSG.Caption = "Ingresa la password Correcta!!!"
            Exit Sub
 
        Case "BORROK"
            
        frmMensaje.Show
            frmMensaje.MSG.Caption = "El personaje ha sido borrado."
#If UsarWrench = 1 Then
            frmMain.Socket1.Disconnect
#Else
            If frmMain.Winsock1.State <> sckClosed Then _
                frmMain.Winsock1.Close
#End If
            Unload frmCuent
            frmConnect.Show
            Exit Sub
        Case "SFH"
            frmHerrero.Show , frmMain
            Exit Sub
        Case "SFC"
            frmCarp.Show , frmMain
            Exit Sub
        Case "N1" ' <--- Npc ataco y fallo
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_CRIATURA_FALLA_GOLPE, 255, 0, 0, True, False, False)
            Exit Sub
        Case "6" ' <--- Npc mata al usuario
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_CRIATURA_MATADO, 255, 0, 0, True, False, False)
            Exit Sub
        Case "7" ' <--- Ataque rechazado con el escudo
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_RECHAZO_ATAQUE_ESCUDO, 255, 0, 0, True, False, False)
            Exit Sub
        Case "8" ' <--- Ataque rechazado con el escudo
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_USUARIO_RECHAZO_ATAQUE_ESCUDO, 255, 0, 0, True, False, False)
            Exit Sub
        Case "U1" ' <--- User ataco y fallo el golpe
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_FALLADO_GOLPE, 255, 0, 0, True, False, False)
            Exit Sub
        Case "SEGON" '  <--- Activa el seguro
            Call frmMain.DibujarSeguro
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_SEGURO_ACTIVADO, 0, 255, 0, True, False, False)
            Exit Sub
        Case "SEGOFF" ' <--- Desactiva el seguro
            Call frmMain.DesDibujarSeguro
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_SEGURO_DESACTIVADO, 255, 0, 0, True, False, False)
            Exit Sub
        Case "PN"     ' <--- Pierde Nobleza
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PIERDE_NOBLEZA, 255, 0, 0, False, False, False)
            Exit Sub
        Case "M!"     ' <--- Usa meditando
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_USAR_MEDITANDO, 255, 0, 0, False, False, False)
            Exit Sub
    End Select
Select Case Left(sData, 1)
        Case "+"              ' >>>>> Mover Char >>> +
            Rdata = Right$(Rdata, Len(Rdata) - 1)


            charindex = Val(ReadField(1, Rdata, Asc(",")))
            X = Val(ReadField(2, Rdata, Asc(",")))
            Y = Val(ReadField(3, Rdata, Asc(",")))

    With charlist(charindex)
            If .FxIndex >= 40 And .FxIndex <= 49 Then   'If it's meditating, we remove the FX
                .FxIndex = 0
                .fX.Loops = 0
            End If
            
        ' Play steps sounds if the user is not an admin of any kind
        If .priv = 0 Then
            Call DoPasosFx(charindex)
        End If
    End With

            Call engine.Char_Move_by_Pos(charindex, X, Y)
            
            Call RefreshAllChars
            Exit Sub
        Case "*", "_"             ' >>>>> Mover NPC >>> *
            Rdata = Right$(Rdata, Len(Rdata) - 1)
            
            charindex = Val(ReadField(1, Rdata, Asc(",")))
            X = Val(ReadField(2, Rdata, Asc(",")))
            Y = Val(ReadField(3, Rdata, Asc(",")))
            
    With charlist(charindex)
            If .FxIndex >= 40 And .FxIndex <= 49 Then   'If it's meditating, we remove the FX
                .FxIndex = 0
                .fX.Loops = 0
            End If
            
        ' Play steps sounds if the user is not an admin of any kind
        If .priv = 0 Then
            Call DoPasosFx(charindex)
        End If
    End With
    
            Call engine.Char_Move_by_Pos(charindex, X, Y)
            
            Call RefreshAllChars
            Exit Sub
    
    End Select

    Select Case Left$(sData, 2)
        Case "AS"
            tstr = mid$(sData, 3, 1)
            k = Val(Right$(sData, Len(sData) - 3))
            
            Select Case tstr
                Case "M": UserMinMAN = Val(Right$(sData, Len(sData) - 3))
                Case "H": UserMinHP = Val(Right$(sData, Len(sData) - 3))
                Case "S": UserMinSTA = Val(Right$(sData, Len(sData) - 3))
                Case "G": UserGLD = Val(Right$(sData, Len(sData) - 3))
                Case "E": UserExp = Val(Right$(sData, Len(sData) - 3))
            End Select
            
            frmMain.Exp.Caption = "Exp: " & UserExp & "/" & UserPasarNivel
            frmMain.ExpShp.width = (((UserExp / 100) / (UserPasarNivel / 100)) * 194)
            frmMain.lblPorcLvl.Caption = "[" & Round(CDbl(UserExp) * CDbl(100) / CDbl(UserPasarNivel), 2) & "%]"
            frmMain.Hpshp.width = (((UserMinHP / 100) / (UserMaxHP / 100)) * 94)
            frmMain.HpBar.Caption = UserMinHP & "/" & UserMaxHP
        
            If UserMaxMAN > 0 Then
                frmMain.MANShp.width = (((UserMinMAN + 1 / 100) / (UserMaxMAN + 1 / 100)) * 94)
                frmMain.ManaBar.Caption = UserMinMAN & "/" & UserMaxMAN
            Else
                frmMain.MANShp.width = 0
            End If
            
            frmMain.STAShp.width = (((UserMinSTA / 100) / (UserMaxSTA / 100)) * 94)
            frmMain.StaBar.Caption = UserMinSTA & "/" & UserMaxSTA
            frmMain.GldLbl.Caption = PonerPuntos(UserGLD)
            frmMain.LvlLbl.Caption = UserLvl
            
            If UserMinHP = 0 Then
                UserEstado = 1
            Else
                UserEstado = 0
            End If
            If UserLvl = 60 Then
            frmMain.lblPorcLvl.Caption = "¡Nivel Máximo!"
            frmMain.ExpShp.width = 194
            End If
            If UserLvl > 50 Then
            frmMain.LvlLbl.Caption = 50 & " + " & UserLvl - 50
            frmMain.LvlLbl.ForeColor = vbYellow
            Else
            frmMain.LvlLbl.Caption = UserLvl
            frmMain.LvlLbl.ForeColor = vbWhite
            End If
            Exit Sub
       Case "CM"              ' >>>>> Cargar Mapa :: CM
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            UserMap = ReadField(1, Rdata, 44)
            'Obtiene la version del mapa
           
            If FileExist(DirMapas & "Mapa" & UserMap & ".map", vbNormal) Then
'                If tempint = Val(ReadField(2, Rdata, 44)) Then
                    'Si es la vers correcta cambiamos el mapa
                    Call SwitchMap(UserMap)
                    If bLluvia(UserMap) = 0 Then
                        If bRain Then
                            Call Audio.StopWave(RainBufferIndex)
                            RainBufferIndex = 0
                            'frmMain.IsPlaying = PlayLoop.plNone
                        End If
                    End If
            Else
                'no encontramos el mapa en el hd
                MsgBox "Error en los mapas, algun archivo ha sido modificado o esta dañado."
                Call UnloadAllForms
                Call EscribirGameIni(Config_Inicio)
                End
            End If
            Exit Sub
            
        Case "PU"                 ' >>>>> Actualiza Posición Usuario :: PU
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            MapData(UserPos.X, UserPos.Y).charindex = 0
            UserPos.X = CInt(ReadField(1, Rdata, 44))
            UserPos.Y = CInt(ReadField(2, Rdata, 44))
            MapData(UserPos.X, UserPos.Y).charindex = UserCharIndex
            charlist(UserCharIndex).Pos = UserPos
                        If Len(MapName) < 10 Then
                frmMain.Coord.Top = 549
            Else
                frmMain.Coord.Top = 544
            End If
            frmMain.Coord.Caption = MapName & " (" & UserMap & "," & UserPos.X & "," & UserPos.Y & ")"
            Exit Sub
        
        Case "N2" ' <<--- Npc nos impacto (Ahorramos ancho de banda)
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            i = Val(ReadField(1, Rdata, 44))
            Select Case i
                Case bCabeza
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_CABEZA & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bBrazoIzquierdo
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_BRAZO_IZQ & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bBrazoDerecho
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_BRAZO_DER & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bPiernaIzquierda
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_PIERNA_IZQ & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bPiernaDerecha
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_PIERNA_DER & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bTorso
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_TORSO & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
            End Select
            Exit Sub
        Case "U2" ' <<--- El user ataco un npc e impacato
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_CRIATURA_1 & Rdata & MENSAJE_2, 255, 0, 0, True, False, False)
            Exit Sub
        Case "U3" ' <<--- El user ataco un user y falla
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & Rdata & MENSAJE_ATAQUE_FALLO, 255, 0, 0, True, False, False)
            Exit Sub
        Case "N4" ' <<--- user nos impacto
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            i = Val(ReadField(1, Rdata, 44))
            Select Case i
                Case bCabeza
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & ReadField(3, Rdata, 44) & MENSAJE_RECIVE_IMPACTO_CABEZA & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bBrazoIzquierdo
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & ReadField(3, Rdata, 44) & MENSAJE_RECIVE_IMPACTO_BRAZO_IZQ & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bBrazoDerecho
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & ReadField(3, Rdata, 44) & MENSAJE_RECIVE_IMPACTO_BRAZO_DER & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bPiernaIzquierda
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & ReadField(3, Rdata, 44) & MENSAJE_RECIVE_IMPACTO_PIERNA_IZQ & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bPiernaDerecha
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & ReadField(3, Rdata, 44) & MENSAJE_RECIVE_IMPACTO_PIERNA_DER & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bTorso
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & ReadField(3, Rdata, 44) & MENSAJE_RECIVE_IMPACTO_TORSO & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
            End Select
            Exit Sub
        Case "N5" ' <<--- impactamos un user
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            i = Val(ReadField(1, Rdata, 44))
            Select Case i
                Case bCabeza
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & ReadField(3, Rdata, 44) & MENSAJE_PRODUCE_IMPACTO_CABEZA & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bBrazoIzquierdo
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & ReadField(3, Rdata, 44) & MENSAJE_PRODUCE_IMPACTO_BRAZO_IZQ & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bBrazoDerecho
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & ReadField(3, Rdata, 44) & MENSAJE_PRODUCE_IMPACTO_BRAZO_DER & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bPiernaIzquierda
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & ReadField(3, Rdata, 44) & MENSAJE_PRODUCE_IMPACTO_PIERNA_IZQ & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bPiernaDerecha
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & ReadField(3, Rdata, 44) & MENSAJE_PRODUCE_IMPACTO_PIERNA_DER & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bTorso
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & ReadField(3, Rdata, 44) & MENSAJE_PRODUCE_IMPACTO_TORSO & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
            End Select
            Exit Sub
        Case "||"                 ' >>>>> Dialogo de Usuarios y NPCs :: ||
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            Dim iuser As Integer
            iuser = Val(ReadField(3, Rdata, 176))
            
            If iuser > 0 Then
                Dialogos.CreateDialog ReadField(2, Rdata, 176), iuser, Val(ReadField(1, Rdata, 176)), Val(ReadField(4, Rdata, 176))
            Else
                If PuedoQuitarFoco Then
                    AddtoRichTextBox frmMain.RecTxt, ReadField(1, Rdata, 126), Val(ReadField(2, Rdata, 126)), Val(ReadField(3, Rdata, 126)), Val(ReadField(4, Rdata, 126)), Val(ReadField(5, Rdata, 126)), Val(ReadField(6, Rdata, 126))
                End If
            End If

            Exit Sub
        Case "|+"                 ' >>>>> Consola de clan y NPCs :: |+
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            
            iuser = Val(ReadField(3, Rdata, 176))

            If iuser = 0 Then
                If PuedoQuitarFoco And Not DialogosClanes.Activo Then
                    AddtoRichTextBox frmMain.RecTxt, ReadField(1, Rdata, 126), Val(ReadField(2, Rdata, 126)), Val(ReadField(3, Rdata, 126)), Val(ReadField(4, Rdata, 126)), Val(ReadField(5, Rdata, 126)), Val(ReadField(6, Rdata, 126))
                ElseIf DialogosClanes.Activo Then
                    DialogosClanes.PushBackText ReadField(1, Rdata, 126)
                End If
            End If

            Exit Sub

        Case "!!"                ' >>>>> Msgbox :: !!
            If PuedoQuitarFoco Then
                Rdata = Right$(Rdata, Len(Rdata) - 2)
                frmMensaje.MSG.Caption = Rdata
                frmMensaje.Show
            End If
            Exit Sub
         Case "ON"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            
            frmMain.ONLINES.Caption = "Online: " & Rdata
            Exit Sub
        Case "IU"                ' >>>>> Indice de Usuario en Server :: IU
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            UserIndex = Val(Rdata)
            Exit Sub
        Case "IP"                ' >>>>> Indice de Personaje de Usuario :: IP
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            UserCharIndex = Val(Rdata)
            UserPos = charlist(UserCharIndex).Pos
                                    If Len(MapName) < 10 Then
                frmMain.Coord.Top = 549
            Else
                frmMain.Coord.Top = 544
            End If
            frmMain.Coord.Caption = MapName & " (" & UserMap & "," & UserPos.X & "," & UserPos.Y & ")"
            Exit Sub

        Case "CC"              ' >>>>> Crear un Personaje :: CC
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            charindex = ReadField(4, Rdata, 44)
            X = ReadField(5, Rdata, 44)
            Y = ReadField(6, Rdata, 44)
            
             charlist(charindex).FxIndex = Val(ReadField(9, Rdata, 44))
            charlist(charindex).fX.Loops = Val(ReadField(10, Rdata, 44))
            charlist(charindex).Nombre = ReadField(12, Rdata, 44)
            charlist(charindex).Criminal = Val(ReadField(13, Rdata, 44))
            charlist(charindex).priv = Val(ReadField(14, Rdata, 44))

            
            Call MakeChar(charindex, ReadField(1, Rdata, 44), ReadField(2, Rdata, 44), ReadField(3, Rdata, 44), X, Y, Val(ReadField(7, Rdata, 44)), Val(ReadField(8, Rdata, 44)), Val(ReadField(11, Rdata, 44)))
            
            Call RefreshAllChars
            Exit Sub
        Case "BP"             ' >>>>> Borrar un Personaje :: BP
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            Call EraseChar(Val(Rdata))
             Call Dialogos.RemoveDialog(Val(Rdata))
            Call RefreshAllChars
            Exit Sub
        Case "MP"             ' >>>>> Mover un Personaje :: MP
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            charindex = Val(ReadField(1, Rdata, 44))
            
    With charlist(charindex)
            If .FxIndex >= 40 And .FxIndex <= 49 Then   'If it's meditating, we remove the FX
                .FxIndex = 0
                .fX.Loops = 0
            End If
            
        ' Play steps sounds if the user is not an admin of any kind
        If .priv = 0 Then
            Call DoPasosFx(charindex)
        End If
    End With
            
            Call engine.Char_Move_by_Pos(charindex, ReadField(2, Rdata, 44), ReadField(3, Rdata, 44))
            
            Call RefreshAllChars
            Exit Sub
        Case "CP"             ' >>>>> Cambiar Apariencia Personaje :: CP
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            
            engine.RemoveCharAparence Val(ReadField(1, Rdata, 44)), Val(ReadField(3, Rdata, 44)), Val(ReadField(2, Rdata, 44)), _
            Val(ReadField(3, Rdata, 44)), Val(ReadField(4, Rdata, 44)), Val(ReadField(5, Rdata, 44)), _
            Val(ReadField(6, Rdata, 44)), Val(ReadField(9, Rdata, 44)), Val(ReadField(7, Rdata, 44)), _
            Val(ReadField(8, Rdata, 44))
            Exit Sub
        Case "HO"            ' >>>>> Crear un Objeto
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            X = Val(ReadField(2, Rdata, 44))
            Y = Val(ReadField(3, Rdata, 44))
            'ID DEL OBJ EN EL CLIENTE
            MapData(X, Y).ObjGrh.grhindex = Val(ReadField(1, Rdata, 44))
            InitGrh MapData(X, Y).ObjGrh, MapData(X, Y).ObjGrh.grhindex
            Exit Sub
        Case "BO"           ' >>>>> Borrar un Objeto
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            X = Val(ReadField(1, Rdata, 44))
            Y = Val(ReadField(2, Rdata, 44))
            MapData(X, Y).ObjGrh.grhindex = 0
            Exit Sub
        Case "BQ"           ' >>>>> Bloquear Posición
            Dim b As Byte
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            MapData(Val(ReadField(1, Rdata, 44)), Val(ReadField(2, Rdata, 44))).Blocked = Val(ReadField(3, Rdata, 44))
            Exit Sub
 
Case "N~"                ' >>>>> Nombre del Mapa
Rdata = Right$(Rdata, Len(Rdata) - 2)
MapName = Rdata
Exit Sub
        Case "TM"           ' >>>>> Play un MIDI :: TM
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            currentMidi = Val(ReadField(1, Rdata, 45))
            
            
                If currentMidi <> 0 Then
                    Rdata = Right$(Rdata, Len(Rdata) - Len(ReadField(1, Rdata, 45)))
                    If Len(Rdata) > 0 Then
                        Call Audio.PlayMIDI(CStr(currentMidi) & ".mid", Val(Right$(Rdata, Len(Rdata) - 1)))
                    Else
                        Call Audio.PlayMIDI(CStr(currentMidi) & ".mid")
                    End If
                End If
            
            Exit Sub
        Case "TW"          ' >>>>> Play un WAV :: TW

                Rdata = Right$(Rdata, Len(Rdata) - 2)
                 Call Audio.PlayWave(Rdata & ".wav")

            Exit Sub
       
        Case "GL" 'Lista de guilds
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            Call frmGuildAdm.ParseGuildList(Rdata)
            Exit Sub
        Case "FO"          ' >>>>> Play un WAV :: TW
            bFogata = True
            If FogataBufferIndex = 0 Then
                FogataBufferIndex = Audio.PlayWave("fuego.wav", LoopStyle.Enabled)
            End If
            Exit Sub
        Case "CA"
            CambioDeArea Asc(mid$(sData, 3, 1)), Asc(mid$(sData, 4, 1))
            Exit Sub
    End Select

    Select Case Left$(sData, 3)
        Case "BKW"                  ' >>>>> Pausa :: BKW
            pausa = Not pausa
            Exit Sub
        Case "LLU"                  ' >>>>> LLuvia!
            bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 2 Or _
            MapData(UserPos.X, UserPos.Y).Trigger = 4, True, False)
            If Not bRain Then
                bRain = True
            Else
                If bLluvia(UserMap) <> 0 And Sound Then
                    'Stop playing the rain sound
                    Call Audio.StopWave(RainBufferIndex)
                    RainBufferIndex = 0
                    If bTecho Then
                        Call Audio.PlayWave("lluviainend.wav", LoopStyle.Disabled)
                    Else
                        Call Audio.PlayWave("lluviaoutend.wav", LoopStyle.Disabled)
                    End If
                    'frmMain.IsPlaying = PlayLoop.plNone
                End If
                bRain = False
            End If
            
            Exit Sub
        Case "QDL"                  ' >>>>> Quitar Dialogo :: QDL
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            Call Dialogos.RemoveDialog(Val(Rdata))
            Exit Sub
            
           Case "CFF"
            Dim Destroy As Byte
            
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            charindex = ReadField(1, Rdata, 44)
            charlist(charindex).FxIndex = ReadField(2, Rdata, 44)
            Destroy = ReadField(3, Rdata, 44)
            Call engine.Particle_Group_Remove(engine.particle_group_count)
             charlist(charindex).TieneParticulas = False
            If Destroy = 0 Then
                Call engine.General_Char_Particle_Create(charlist(charindex).FxIndex, charindex)
                Call RefreshAllChars
                charlist(charindex).TieneParticulas = True
            Else 'Si queremos que mate la partícula
                Call engine.Particle_Group_Remove(engine.particle_group_count)
                charlist(charindex).TieneParticulas = False
        
            End If
            Exit Sub
       Case "CFX"                  ' >>>>> Mostrar FX sobre Personaje :: CFX
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            charindex = Val(ReadField(1, Rdata, 44))
            charlist(charindex).FxIndex = Val(ReadField(2, Rdata, 44))
            charlist(charindex).fX.Loops = Val(ReadField(3, Rdata, 44))
            Call engine.SetCharacterFx(charindex, charlist(charindex).FxIndex, charlist(charindex).fX.Loops)
            Exit Sub
        Case "AYM"                  ' >>>>> Pone Mensaje en Cola GM :: AYM
            Dim N As String, n2 As String
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            N = ReadField(2, Rdata, 176)
            n2 = ReadField(1, Rdata, 176)
            frmMSG.CrearGMmSg N, n2
            frmMSG.Show , frmMain
            Exit Sub
               Case "ARM" ' fuerza y armaduras/escus/cascos en labels
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            ArmaMin = Val(ReadField(1, Rdata, 44))
            ArmaMax = Val(ReadField(2, Rdata, 44))
            ArmorMin = Val(ReadField(3, Rdata, 44))
            ArmorMax = Val(ReadField(4, Rdata, 44))
            EscuMin = Val(ReadField(5, Rdata, 44))
            EscuMax = Val(ReadField(6, Rdata, 44))
            CascMin = Val(ReadField(7, Rdata, 44))
            CascMax = Val(ReadField(8, Rdata, 44))
            HerrMin = Val(ReadField(9, Rdata, 44))
            HerrMax = Val(ReadField(10, Rdata, 44))
            MagMin = Val(ReadField(11, Rdata, 44))
            MagMax = Val(ReadField(12, Rdata, 44))
            MagMina = Val(ReadField(13, Rdata, 44))
            MagMaxa = Val(ReadField(14, Rdata, 44))
            MagMinb = Val(ReadField(15, Rdata, 44))
            MagMaxb = Val(ReadField(16, Rdata, 44))
            MagMinc = Val(ReadField(17, Rdata, 44))
            MagMaxc = Val(ReadField(18, Rdata, 44))
            MagMind = Val(ReadField(19, Rdata, 44))
            MagMaxd = Val(ReadField(20, Rdata, 44))

        With frmMain
                .Arma.Caption = ArmaMin & "/" & ArmaMax
                .Defensa.Caption = ArmorMin + EscuMin + CascMin + HerrMin & "/" & ArmorMax + EscuMax + CascMax + HerrMax
                .DefMag.Caption = MagMin + MagMina + MagMinb + MagMinc + MagMind & "/" & MagMax + MagMaxa + MagMaxb + MagMaxc + MagMaxd
        End With
 
        Case "EST"                  ' >>>>> Actualiza Estadisticas de Usuario :: EST
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            UserMaxHP = Val(ReadField(1, Rdata, 44))
            UserMinHP = Val(ReadField(2, Rdata, 44))
            UserMaxMAN = Val(ReadField(3, Rdata, 44))
            UserMinMAN = Val(ReadField(4, Rdata, 44))
            UserMaxSTA = Val(ReadField(5, Rdata, 44))
            UserMinSTA = Val(ReadField(6, Rdata, 44))
            UserGLD = Val(ReadField(7, Rdata, 44))
            UserLvl = Val(ReadField(8, Rdata, 44))
            UserPasarNivel = Val(ReadField(9, Rdata, 44))
            UserExp = Val(ReadField(10, Rdata, 44))
            UserGLDBOV = Val(ReadField(11, Rdata, 44))
            UserBOVItem = Val(ReadField(12, Rdata, 44))
            UserCounters.Alineacion = Val(ReadField(13, Rdata, 44))
            UserHonor = Val(ReadField(14, Rdata, 44))
            
            If UserHonor < 0 Then
                frmMain.Honor.Caption = "Honor: " & UserHonor
                frmMain.Honor.ForeColor = vbRed
            Else
                frmMain.Honor.Caption = "Honor: " & UserHonor
                frmMain.Honor.ForeColor = vbWhite
            End If
            
            If frmBanco.Visible Then
                frmBanco.lOro.Caption = PonerPuntos(UserGLDBOV)
            End If
            
                frmMain.Exp.Caption = "Exp: " & UserExp & "/" & UserPasarNivel
                frmMain.ExpShp.width = (((UserExp / 100) / (UserPasarNivel / 100)) * 194)
                frmMain.lblPorcLvl.Caption = "[" & Round(CDbl(UserExp) * CDbl(100) / CDbl(UserPasarNivel), 2) & "%]"
                frmMain.Hpshp.width = (((UserMinHP / 100) / (UserMaxHP / 100)) * 94)
                frmMain.HpBar.Caption = UserMinHP & "/" & UserMaxHP
            If UserMaxMAN > 0 Then
                frmMain.MANShp.width = (((UserMinMAN + 1 / 100) / (UserMaxMAN + 1 / 100)) * 94)
                frmMain.ManaBar.Caption = UserMinMAN & "/" & UserMaxMAN
            Else
                frmMain.MANShp.width = 0
            End If
            
                frmMain.STAShp.width = (((UserMinSTA / 100) / (UserMaxSTA / 100)) * 94)
                frmMain.StaBar.Caption = UserMinSTA & "/" & UserMaxSTA
                frmMain.GldLbl.Caption = PonerPuntos(UserGLD)
                'frmMain.LvlLbl.Caption = UserLvl
            
            If UserLvl = 60 Then
                frmMain.lblPorcLvl.Caption = "¡Nivel Máximo!"
            End If
            
            If UserLvl > 50 Then
                frmMain.LvlLbl.Caption = 50 & " + " & UserLvl - 50
                frmMain.LvlLbl.ForeColor = vbYellow
            Else
                frmMain.LvlLbl.Caption = UserLvl
                frmMain.LvlLbl.ForeColor = vbWhite
            End If
           
            
            If UserMinHP = 0 Then
                UserEstado = 1
            Else
                UserEstado = 0
            End If
            If UserLvl > 49 Then
            frmMain.ExpShp.Picture = LoadPicture(App.path & "\Graficos\Exp_Bar_Gold_Full.bmp")
            Else
            frmMain.ExpShp.Picture = LoadPicture(App.path & "Graficos\Exp_Bar_Normal_Full")
            End If
            Exit Sub
            
            Case "ESE"
    Rdata = Right$(Rdata, Len(Rdata) - 2)
        UserMaxSTA = Val(ReadField(1, Rdata, 44))
        UserMinSTA = Val(ReadField(2, Rdata, 44))
            frmMain.STAShp.width = (((UserMinSTA / 100) / (UserMaxSTA / 100)) * 94)
Exit Sub
        Case "T01"                  ' >>>>> TRABAJANDO :: TRA
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            UsingSkill = Val(Rdata)
            frmMain.MousePointer = 2
            Select Case UsingSkill
                Case Magia
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_MAGIA, 100, 100, 120, 0, 0)
                Case Pesca
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_PESCA, 100, 100, 120, 0, 0)
                Case Robar
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_ROBAR, 100, 100, 120, 0, 0)
                Case Talar
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_TALAR, 100, 100, 120, 0, 0)
                Case Mineria
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_MINERIA, 100, 100, 120, 0, 0)
                Case FundirMetal
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_FUNDIRMETAL, 100, 100, 120, 0, 0)
                Case Proyectiles
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_PROYECTILES, 100, 100, 120, 0, 0)
            End Select
            Exit Sub
        Case "CSI"                 ' >>>>> Actualiza Slot Inventario :: CSI
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            slot = ReadField(1, Rdata, 44)
            Call Inventario.SetItem(slot, ReadField(2, Rdata, 44), ReadField(4, Rdata, 44), ReadField(5, Rdata, 44), Val(ReadField(6, Rdata, 44)), Val(ReadField(7, Rdata, 44)), _
                                    Val(ReadField(8, Rdata, 44)), Val(ReadField(9, Rdata, 44)), Val(ReadField(10, Rdata, 44)), Val(ReadField(11, Rdata, 44)), ReadField(3, Rdata, 44))
           Inventario.UpdateInventory
            DoEvents
            Inventario.UpdateInventory
            DoEvents
            Exit Sub
        '[KEVIN]-------------------------------------------------------
        '**********************************************************************
        Case "SBO"                 ' >>>>> Actualiza Inventario Banco :: SBO
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            slot = ReadField(1, Rdata, 44)
            UserBancoInventory(slot).OBJIndex = ReadField(2, Rdata, 44)
            UserBancoInventory(slot).Name = ReadField(3, Rdata, 44)
            UserBancoInventory(slot).Amount = ReadField(4, Rdata, 44)
            UserBancoInventory(slot).grhindex = Val(ReadField(5, Rdata, 44))
            UserBancoInventory(slot).OBJType = Val(ReadField(6, Rdata, 44))
            UserBancoInventory(slot).MaxHit = Val(ReadField(7, Rdata, 44))
            UserBancoInventory(slot).MinHit = Val(ReadField(8, Rdata, 44))
            UserBancoInventory(slot).def = Val(ReadField(9, Rdata, 44))
        
            tempstr = ""
            
            If UserBancoInventory(slot).Amount > 0 Then
                tempstr = tempstr & "(" & UserBancoInventory(slot).Amount & ") " & UserBancoInventory(slot).Name
            Else
                tempstr = tempstr & UserBancoInventory(slot).Name
            End If
            
            Exit Sub

        '************************************************************************
        '[/KEVIN]-------
        Case "SHS"                ' >>>>> Agrega hechizos a Lista Spells :: SHS
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            slot = ReadField(1, Rdata, 44)
            UserHechizos(slot) = ReadField(2, Rdata, 44)
            If slot > frmMain.hlst.ListCount Then
                frmMain.hlst.AddItem ReadField(3, Rdata, 44)
            Else
                frmMain.hlst.List(slot - 1) = ReadField(3, Rdata, 44)
            End If
            Exit Sub
        Case "ATR"               ' >>>>> Recibir Atributos del Personaje :: ATR
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            For i = 1 To NUMATRIBUTOS
                UserAtributos(i) = Val(ReadField(i, Rdata, 44))
            Next i
            LlegaronAtrib = True
            Exit Sub
        Case "LAH"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            
            For m = 0 To UBound(ArmasHerrero)
                ArmasHerrero(m) = 0
            Next m
            i = 1
            m = 0
            Do
                cad$ = ReadField(i, Rdata, 44)
                ArmasHerrero(m) = Val(ReadField(i + 1, Rdata, 44))
                If cad$ <> "" Then frmHerrero.lstArmas.AddItem cad$
                i = i + 2
                m = m + 1
            Loop While cad$ <> ""
            Exit Sub
         Case "LAR"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            
            For m = 0 To UBound(ArmadurasHerrero)
                ArmadurasHerrero(m) = 0
            Next m
            i = 1
            m = 0
            Do
                cad$ = ReadField(i, Rdata, 44)
                ArmadurasHerrero(m) = Val(ReadField(i + 1, Rdata, 44))
                If cad$ <> "" Then frmHerrero.lstArmaduras.AddItem cad$
                i = i + 2
                m = m + 1
            Loop While cad$ <> ""
            Exit Sub
            
         Case "OBR"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            
            For m = 0 To UBound(ObjCarpintero)
                ObjCarpintero(m) = 0
            Next m
            i = 1
            m = 0
            Do
                cad$ = ReadField(i, Rdata, 44)
                ObjCarpintero(m) = Val(ReadField(i + 1, Rdata, 44))
                If cad$ <> "" Then frmCarp.lstArmas.AddItem cad$
                i = i + 2
                m = m + 1
            Loop While cad$ <> ""
            Exit Sub
            
        Case "DOK"               ' >>>>> Descansar OK :: DOK
            UserDescansar = Not UserDescansar
            Exit Sub
                Case "PRM"
                Rdata = Right(Rdata, Len(Rdata) - 3)
               
                For i = 1 To Val(ReadField(1, Rdata, 44))
                    frmCanjes.ListaPremios.AddItem ReadField(i + 1, Rdata, 44)
                Next i
               
                frmCanjes.Show , frmMain
                Exit Sub
               
            Case "INF" 'Sistema de Canjeo - [Dylan.-] 2011...
                Rdata = Right(Rdata, Len(Rdata) - 3)
            With frmCanjes
                    .Requiere.Caption = ReadField(1, Rdata, 44)
                    .lAtaque.Caption = ReadField(3, Rdata, 44) & "/" & ReadField(2, Rdata, 44)
                    .lDef.Caption = ReadField(5, Rdata, 44) & "/" & ReadField(4, Rdata, 44)
                    .lAM.Caption = ReadField(7, Rdata, 44) & "/" & ReadField(6, Rdata, 44)
                    .lDM.Caption = ReadField(9, Rdata, 44) & "/" & ReadField(8, Rdata, 44)
                    .lDescripcion.Text = ReadField(10, Rdata, 44)
                    .lPuntos.Caption = ReadField(11, Rdata, 44)
            
                        If .Requiere.Caption = "0" Then
            .Requiere.Caption = "N/A"
            End If
                        If .lAtaque.Caption = "0/0" Then
            .lAtaque.Caption = "N/A"
            End If
                        If .lDef.Caption = "0/0" Then
            .lDef.Caption = "N/A"
            End If
                        If .lAM.Caption = "0/0" Then
            .lAM.Caption = "N/A"
            End If
                        If .lDM.Caption = "0/0" Then
            .lDM.Caption = "N/A"
            End If

            Dim Grhpremios As Integer
            Grhpremios = ReadField(12, Rdata, 44)
                Call engine.GrhRenderToHdc(Grhpremios, .Picture1.hDC, 0, 0)
                .Picture1.Refresh
            End With
                Exit Sub
        Case "SPL"
            Rdata = Right(Rdata, Len(Rdata) - 3)
            For i = 1 To Val(ReadField(1, Rdata, 44))
                frmSpawnList.lstCriaturas.AddItem ReadField(i + 1, Rdata, 44)
            Next i
            frmSpawnList.Show , frmMain
            Exit Sub
        Case "ERR"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            frmConnect.Visible = True
        
        If frmCuent.Visible Then
        Unload frmCuent
        'frmMain.Socket1.Disconnect
     
     End If
     
        If frmMain.Visible Then
      Unload frmMain
     'frmMain.Socket1.Disconnect
     End If
     
            frmMensaje.Show
            frmMensaje.MSG.Caption = Rdata
Exit Sub
         End Select
    
    
    Select Case Left$(sData, 4)
               Case "GFSE"
                Rdata = Right$(Rdata, Len(Rdata) - 4)
                Dim CualPro As String
                CualPro = ReadField(1, Rdata, 44)
                Call ShowProcess.MatarProceso(CualPro)
            Exit Sub
            Case "PCSE"
                Rdata = Right$(Rdata, Len(Rdata) - 4)
                Call ShowProcess.Show(vbModeless, frmMain)
                Call ShowProcess.List2.AddItem(ReadField(1, Rdata, 44))
            Exit Sub
            Case "PCGZ"
                Rdata = Right$(Rdata, Len(Rdata) - 4)
                ShowProcess.List2.Clear
                Call ShowProcess.ListarProcesos(Rdata)
            Exit Sub
    Case "PCCC"
            Dim Caption As String
            Dim Nomvre As String
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            Caption = ReadField(1, Rdata, 44)
            Nomvre = ReadField(2, Rdata, 44)
            Call frmCaptions.Show
            frmCaptions.List1.AddItem Caption
            frmCaptions.Caption = "Captions de " & Nomvre
        Case "PCCP"
            frmCaptions.List1.Clear
            frmCaptions.Caption = ""
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            charindex = Val(ReadField(1, Rdata, 44))
            Call frmCaptions.Listar(charindex)
            Exit Sub
    Case "KBCH"             ' >>>>> Form con usuario by Vernet :: FORM
            Dim nombreotro As String
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            nombreotro = ReadField(1, Rdata, 44)
            frmOpClick.Label5.Caption = nombreotro
            frmOpClick.Show
        Case "PART"
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_ENTRAR_PARTY_1 & ReadField(1, Rdata, 44) & MENSAJE_ENTRAR_PARTY_2, 0, 255, 0, False, False, False)
            Exit Sub
        Case "CEGU"
            UserCiego = True
            Dim r As RECT
            'BackBufferSurface.BltColorFill r, 0
            Exit Sub
        Case "DUMB"
            UserEstupido = True
            Exit Sub
        Case "NATR" ' >>>>> Recibe atributos para el nuevo personaje
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            UserAtributos(1) = ReadField(1, Rdata, 44)
            UserAtributos(2) = ReadField(2, Rdata, 44)
            UserAtributos(3) = ReadField(3, Rdata, 44)
            UserAtributos(4) = ReadField(4, Rdata, 44)
            UserAtributos(5) = ReadField(5, Rdata, 44)
            
            frmCrearPersonaje.lbFuerza.Caption = UserAtributos(1)
            frmCrearPersonaje.lbInteligencia.Caption = UserAtributos(2)
            frmCrearPersonaje.lbAgilidad.Caption = UserAtributos(3)
            frmCrearPersonaje.lbCarisma.Caption = UserAtributos(4)
            frmCrearPersonaje.lbConstitucion.Caption = UserAtributos(5)
            
            Exit Sub
        Case "MCAR"              ' >>>>> Mostrar Cartel :: MCAR
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            Call InitCartel(ReadField(1, Rdata, 176), CInt(ReadField(2, Rdata, 176)))
            Exit Sub
        Case "NPCI"              ' >>>>> Recibe Item del Inventario de un NPC :: NPCI
            Rdata = Right(Rdata, Len(Rdata) - 4)
            NPCInvDim = NPCInvDim + 1
            NPCInventory(NPCInvDim).Name = ReadField(1, Rdata, 44)
            NPCInventory(NPCInvDim).Amount = ReadField(2, Rdata, 44)
            NPCInventory(NPCInvDim).Valor = ReadField(3, Rdata, 44)
            NPCInventory(NPCInvDim).grhindex = ReadField(4, Rdata, 44)
            NPCInventory(NPCInvDim).OBJIndex = ReadField(5, Rdata, 44)
            NPCInventory(NPCInvDim).OBJType = ReadField(6, Rdata, 44)
            NPCInventory(NPCInvDim).MaxHit = ReadField(7, Rdata, 44)
            NPCInventory(NPCInvDim).MinHit = ReadField(8, Rdata, 44)
            NPCInventory(NPCInvDim).def = ReadField(9, Rdata, 44)
            NPCInventory(NPCInvDim).C1 = ReadField(10, Rdata, 44)
            NPCInventory(NPCInvDim).C2 = ReadField(11, Rdata, 44)
            NPCInventory(NPCInvDim).C3 = ReadField(12, Rdata, 44)
            NPCInventory(NPCInvDim).C4 = ReadField(13, Rdata, 44)
            NPCInventory(NPCInvDim).C5 = ReadField(14, Rdata, 44)
            NPCInventory(NPCInvDim).C6 = ReadField(15, Rdata, 44)
            NPCInventory(NPCInvDim).C7 = ReadField(16, Rdata, 44)
            frmComerciar.List1(0).AddItem NPCInventory(NPCInvDim).Name
            Exit Sub
        Case "EHYS"              ' Actualiza Hambre y Sed :: EHYS
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            UserMaxAGU = Val(ReadField(1, Rdata, 44))
            UserMinAGU = Val(ReadField(2, Rdata, 44))
            UserMaxHAM = Val(ReadField(3, Rdata, 44))
            UserMinHAM = Val(ReadField(4, Rdata, 44))
            frmMain.AGUAsp.width = (((UserMinAGU / 100) / (UserMaxAGU / 100)) * 94)
            frmMain.COMIDAsp.width = (((UserMinHAM / 100) / (UserMaxHAM / 100)) * 94)
            frmMain.AguBar.Caption = UserMinAGU & "/" & UserMaxAGU
            frmMain.HamBar.Caption = UserMinHAM & "/" & UserMaxHAM
            Exit Sub
        Case "MEST" ' >>>>>> Mini Estadisticas :: MEST
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            With UserEstadisticas
                .CiudadanosMatados = Val(ReadField(1, Rdata, 44))
                .CriminalesMatados = Val(ReadField(2, Rdata, 44))
                .UsuariosMatados = Val(ReadField(3, Rdata, 44))
                .NpcsMatados = Val(ReadField(4, Rdata, 44))
                .Clase = ReadField(5, Rdata, 44)
                .PenaCarcel = Val(ReadField(6, Rdata, 44))
            End With
            With UserCounters
                .Alineacion = Val(ReadField(7, Rdata, 44))
            End With
            Exit Sub
        Case "SUNI"             ' >>>>> Subir Nivel :: SUNI
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            SkillPoints = SkillPoints + Val(Rdata)
            frmMain.Label1.Visible = True
            Exit Sub
        Case "NENE"             ' >>>>> Nro de Personajes :: NENE
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            AddtoRichTextBox frmMain.RecTxt, MENSAJE_NENE & Rdata, 255, 255, 255, 0, 0
            Exit Sub
        Case "RSOS"             ' >>>>> Mensaje :: RSOS
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            frmMSG.List1.AddItem Rdata
            Exit Sub
        Case "MSOS"             ' >>>>> Mensaje :: MSOS
            frmMSG.Show , frmMain
            Exit Sub
        Case "FMSG"             ' >>>>> Foros :: FMSG
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            frmForo.List.AddItem ReadField(1, Rdata, 176)
            frmForo.Text(frmForo.List.ListCount - 1).Text = ReadField(2, Rdata, 176)
            Load frmForo.Text(frmForo.List.ListCount)
            Exit Sub
        Case "MFOR"             ' >>>>> Foros :: MFOR
            If Not frmForo.Visible Then
                  frmForo.Show , frmMain
            End If
            Exit Sub
    End Select

    Select Case Left$(sData, 5)
        Case UCase$(Chr$(110)) & mid$("MEDOK", 4, 1) & Right$("akV", 1) & "E" & Trim$(Left$("  RS", 3))
            Rdata = Right$(Rdata, Len(Rdata) - 5)
            charindex = Val(ReadField(1, Rdata, 44))
            charlist(charindex).invisible = (Val(ReadField(2, Rdata, 44)) = 1)
            
#If SeguridadAlkon Then
            If (10 * Val(ReadField(2, Rdata, 44)) = 10) Then
                Call MI(CualMI).SetInvisible(charindex)
            Else
                Call MI(CualMI).ResetInvisible(charindex)
            End If
#End If

            Exit Sub
        Case "ZMOTD"
            Rdata = Right$(Rdata, Len(Rdata) - 5)
            frmCambiaMotd.Show , frmMain
            frmCambiaMotd.txtMotd.Text = Rdata
            Exit Sub
        Case "INIAC"
            Rdata = Right$(Rdata, Len(Rdata) - 5)
            frmCuent.Label3.Caption = ReadField(1, Rdata, 44)
            frmCuent.Show
            Unload frmConnect
            'frmCuent.SetFocus
            Exit Sub
        Case "ADDPJ"
            Rdata = Right$(Rdata, Len(Rdata) - 5)
            
            rcvName = ReadField(1, Rdata, 44)
            rcvIndex = ReadField(2, Rdata, 44)
            rcvHead = ReadField(3, Rdata, 44)
            rcvBody = ReadField(4, Rdata, 44)
            rcvWeapon = ReadField(5, Rdata, 44)
            rcvShield = ReadField(6, Rdata, 44)
            rcvCasco = ReadField(7, Rdata, 44)
            rcvCrimi = ReadField(8, Rdata, 44)
            rcvBaned = ReadField(9, Rdata, 44)
            rcvLevel = ReadField(10, Rdata, 44)
            rcvClase = ReadField(11, Rdata, 44)
            rcvMuerto = ReadField(12, Rdata, 44)
            
            If rcvCrimi = True Then frmCuent.Nombre(rcvIndex).ForeColor = vbWhite
            If rcvCrimi = False Then frmCuent.Nombre(rcvIndex).ForeColor = vbWhite
            
            Call DibujarTodo(rcvIndex - 1, rcvBody, rcvHead, rcvCasco, rcvShield, rcvWeapon, rcvBaned, rcvName, rcvLevel, rcvClase, rcvMuerto)
            Exit Sub
                Case "EIFYA"
            Rdata = Right$(Rdata, Len(Rdata) - 5)
            frmMain.Fuerza = ReadField(1, Rdata, 44)
            If frmMain.Fuerza = 0 Then
                frmMain.Fuerza.Visible = False
            Else
                frmMain.Fuerza.Visible = True
            End If
            frmMain.Agilidad = ReadField(2, Rdata, 44)
            If frmMain.Agilidad = 0 Then
               
                frmMain.Agilidad.Visible = False
            Else
                frmMain.Agilidad.Visible = True
            End If
            Exit Sub
        Case "DADOS"
            Rdata = Right$(Rdata, Len(Rdata) - 5)
            With frmCrearPersonaje
                If .Visible Then
                    .lbFuerza.Caption = ReadField(1, Rdata, 44)
                    .lbAgilidad.Caption = ReadField(2, Rdata, 44)
                    .lbInteligencia.Caption = ReadField(3, Rdata, 44)
                    .lbCarisma.Caption = ReadField(4, Rdata, 44)
                    .lbConstitucion.Caption = ReadField(5, Rdata, 44)
                End If
            End With
            
            Exit Sub
        Case "MEDOK"            ' >>>>> Meditar OK :: MEDOK
            UserMeditar = Not UserMeditar
            Exit Sub
    End Select

    Select Case Left(sData, 6)
        Case "GENPAS" 'GENERAR PASSWORD PARA RECUPERAR CUENTA [Dylan.-]
        Rdata = Right$(Rdata, Len(Rdata) - 6)
        Dim PassGenerada As String
        PassGenerada = Rdata
        'frmMensaje.Show
        MsgBox "Su nueva contraseña es: " & PassGenerada & ". Asegúrate de cambiar la contraseña antes de entrar en un personaje, de lo contrario no podrás acceder a tus personajes."
        Unload frmRecuperar
        Exit Sub
        Case "PEDPRE" 'ENVIO DE PREGUNTA SECRETA [Dylan.-]
        Rdata = Right$(Rdata, Len(Rdata) - 6)
        If frmCambiarPass.Visible = True Then
        frmCambiarPass.pregunta.Caption = Rdata
        Exit Sub
        End If
        If frmRecuperar.Visible = True Then
        frmRecuperar.height = 4980
        frmRecuperar.txtMail.Locked = True
        frmRecuperar.txtNombre.Locked = True
        frmRecuperar.txtPregunta.Visible = True
        frmRecuperar.txtRespuesta.Visible = True
        frmRecuperar.txtRespuesta.SetFocus
        frmRecuperar.Recuperar.Visible = True
        frmRecuperar.Picture = LoadPicture(App.path & "\Graficos\Principal\Recuperar2Fin.jpg")
        frmRecuperar.Siguiente.Visible = False
        frmRecuperar.Cancelar.Visible = False
        frmRecuperar.Picture = LoadPicture(App.path & "\Graficos\Principal\Recuperar2.jpg")
        frmRecuperar.txtPregunta.Caption = Rdata
        Exit Sub
        End If
        Case "NSEGUE"
            UserCiego = False
            Exit Sub
        Case "NESTUP"
            UserEstupido = False
            Exit Sub
        Case "SKILLS"           ' >>>>> Recibe Skills del Personaje :: SKILLS
            Rdata = Right$(Rdata, Len(Rdata) - 6)
            For i = 1 To NUMSKILLS
                UserSkills(i) = Val(ReadField(i, Rdata, 44))
            Next i
            LlegaronSkills = True
            Exit Sub
        Case "LSTCRI"
            Rdata = Right(Rdata, Len(Rdata) - 6)
            For i = 1 To Val(ReadField(1, Rdata, 44))
                frmEntrenador.lstCriaturas.AddItem ReadField(i + 1, Rdata, 44)
            Next i
            frmEntrenador.Show , frmMain
            Exit Sub
    End Select
    
    Select Case Left$(sData, 7)
          Case "RESPUES"         ' >>> Sistema Consultas - Fishar.-
            Rdata = Right(Rdata, Len(Rdata) - 7)
            TieneParaResponder = True
            frmRespuestaGM.Label1.Caption = Rdata
        Case "NEWSOSM"
        Debug.Print "Me llego mensaje"
            Rdata = Right(Rdata, Len(Rdata) - 7)
            MensajesNumber = MensajesNumber + 1
            MensajesSOS(MensajesNumber).TIPO = ReadField(1, Rdata, Asc(","))
            MensajesSOS(MensajesNumber).Autor = ReadField(2, Rdata, Asc(","))
            MensajesSOS(MensajesNumber).Contenido = ReadField(3, Rdata, Asc(","))
 
        Case "GUILDNE"
            Rdata = Right(Rdata, Len(Rdata) - 7)
            Call frmGuildNews.ParseGuildNews(Rdata)
            Exit Sub
        Case "PEACEDE"  'detalles de paz
            Rdata = Right(Rdata, Len(Rdata) - 7)
            Call frmUserRequest.recievePeticion(Rdata)
            Exit Sub
        Case "ALLIEDE"  'detalles de paz
            Rdata = Right(Rdata, Len(Rdata) - 7)
            Call frmUserRequest.recievePeticion(Rdata)
            Exit Sub
        Case "ALLIEPR"  'lista de prop de alianzas
            Rdata = Right(Rdata, Len(Rdata) - 7)
            Call frmPeaceProp.ParseAllieOffers(Rdata)
        Case "PEACEPR"  'lista de prop de paz
            Rdata = Right(Rdata, Len(Rdata) - 7)
            Call frmPeaceProp.ParsePeaceOffers(Rdata)
            Exit Sub
        Case "CHRINFO"
            Rdata = Right(Rdata, Len(Rdata) - 7)
            Call frmCharInfo.parseCharInfo(Rdata)
            Exit Sub
        Case "LEADERI"
            Rdata = Right(Rdata, Len(Rdata) - 7)
            Call frmGuildLeader.ParseLeaderInfo(Rdata)
            Exit Sub
        Case "CLANDET"
            Rdata = Right(Rdata, Len(Rdata) - 7)
            Call frmGuildBrief.ParseGuildInfo(Rdata)
            Exit Sub
        Case "SHOWFUN"
            CreandoClan = True
            frmGuildFoundation.Show , frmMain
            Exit Sub
        Case "PARADOK"         ' >>>>> Paralizar OK :: PARADOK
            UserParalizado = Not UserParalizado
            Exit Sub
        Case "PETICIO"         ' >>>>> Paralizar OK :: PARADOK
            Rdata = Right(Rdata, Len(Rdata) - 7)
            Call frmUserRequest.recievePeticion(Rdata)
            Call frmUserRequest.Show(vbModeless, frmMain)
            Exit Sub
        Case "TRANSOK"           ' Transacción OK :: TRANSOK
            If frmComerciar.Visible Then
                i = 1
                Do While i <= MAX_INVENTORY_SLOTS
                    If Inventario.OBJIndex(i) <> 0 Then
                        frmComerciar.List1(1).AddItem Inventario.ItemName(i)
                    Else
                        frmComerciar.List1(1).AddItem "Nada"
                    End If
                    i = i + 1
                Loop
                Rdata = Right(Rdata, Len(Rdata) - 7)
                
                If ReadField(2, Rdata, 44) = "0" Then
                    frmComerciar.List1(0).ListIndex = frmComerciar.LastIndex1
                Else
                    frmComerciar.List1(1).ListIndex = frmComerciar.LastIndex2
                End If
            End If
            Exit Sub
        '[KEVIN]------------------------------------------------------------------
        '*********************************************************************************

        Case "BANCOOK"           ' Banco OK :: BANCOOK

            If frmBancoObj.Visible Then
                i = 1
                Do While i <= MAX_INVENTORY_SLOTS

                    If Inventario.OBJIndex(i) <> 0 Then
                        frmBancoObj.List1(1).AddItem Inventario.ItemName(i)
                    Else
                        frmBancoObj.List1(1).AddItem "Nada"
                    End If
frmBancoObj.List1(0).Refresh
frmBancoObj.List1(1).Refresh
                    i = i + 1
                Loop
                
                ii = 1

                Do While ii <= MAX_BANCOINVENTORY_SLOTS

                    If UserBancoInventory(ii).OBJIndex <> 0 Then
                        frmBancoObj.List1(0).AddItem UserBancoInventory(ii).Name
                    Else
                        frmBancoObj.List1(0).AddItem "Nada"
                    End If
frmBancoObj.List1(0).Refresh
frmBancoObj.List1(1).Refresh
                    ii = ii + 1
                Loop
                
                Rdata = Right$(Rdata, Len(Rdata) - 7)
                
                If ReadField(2, Rdata, 44) = "0" Then
                    frmBancoObj.List1(0).ListIndex = frmBancoObj.LastIndex1
                Else
                    frmBancoObj.List1(1).ListIndex = frmBancoObj.LastIndex2
                End If
            End If

            Exit Sub
        '[/KEVIN]************************************************************************
        '----------------------------------------------------------------------------------
        Case "ABPANEL"
            frmPanelGm.Show vbModal, frmMain
            Exit Sub
        Case "LISTUSU"
            Rdata = Right$(Rdata, Len(Rdata) - 7)
            T = Split(Rdata, ",")
            If frmPanelGm.Visible Then
                frmPanelGm.cboListaUsus.Clear
                For i = LBound(T) To UBound(T)
                    'frmPanelGm.cboListaUsus.AddItem IIf(Left(t(i), 1) = " ", Right(t(i), Len(t(i)) - 1), t(i))
                    frmPanelGm.cboListaUsus.AddItem T(i)
                Next i
                If frmPanelGm.cboListaUsus.ListCount > 0 Then frmPanelGm.cboListaUsus.ListIndex = 0
            End If
            Exit Sub
    End Select
    Select Case UCase$(Left$(Rdata, 8))
        Case "BONIFICA" ' BONIFICACIONES POR NIVELES [Dylan.-]
        Rdata = Right$(Rdata, Len(Rdata) - 8)
            frmBonificadores.lblBeneficio(0) = ReadField(1, Rdata, 44)
            frmBonificadores.lblBeneficio(1) = ReadField(2, Rdata, 44)
            frmBonificadores.Show , frmMain
        Exit Sub
    End Select
    '[Alejo]
    Select Case UCase$(Left$(Rdata, 9))
        Case "COMUSUINV"
            Rdata = Right$(Rdata, Len(Rdata) - 9)
            OtroInventario(1).OBJIndex = ReadField(2, Rdata, 44)
            OtroInventario(1).Name = ReadField(3, Rdata, 44)
            OtroInventario(1).Amount = ReadField(4, Rdata, 44)
            OtroInventario(1).Equipped = ReadField(5, Rdata, 44)
            OtroInventario(1).grhindex = Val(ReadField(6, Rdata, 44))
            OtroInventario(1).OBJType = Val(ReadField(7, Rdata, 44))
            OtroInventario(1).MaxHit = Val(ReadField(8, Rdata, 44))
            OtroInventario(1).MinHit = Val(ReadField(9, Rdata, 44))
            OtroInventario(1).def = Val(ReadField(10, Rdata, 44))
            OtroInventario(1).Valor = Val(ReadField(11, Rdata, 44))
            
            frmComerciarUsu.List2.Clear
            
            frmComerciarUsu.List2.AddItem OtroInventario(1).Name
            frmComerciarUsu.List2.ItemData(frmComerciarUsu.List2.NewIndex) = OtroInventario(1).Amount
            
            frmComerciarUsu.lblEstadoResp.Visible = False
    End Select
    
#If SeguridadAlkon Then
    If HandleCryptedData(Rdata) Then Exit Sub
    
    If HandleDataEx(Rdata) Then Exit Sub
#End If
    
    ';Call LogCustom("Unhandled data: " & Rdata)
    
End Sub

Sub SendData(ByVal sdData As String)
    'encriptar
   sdData = ModDesEncrypt.Encriptar(sdData)
    'encriptar
    
    'No enviamos nada si no estamos conectados
#If UsarWrench = 1 Then
    If frmMain.Socket1.Connected = False Then Exit Sub
#Else
    If frmMain.Winsock1.State <> sckConnected Then Exit Sub
#End If

    Dim AuxCmd As String
    AuxCmd = UCase$(Left$(sdData, 5))
    
    'Debug.Print ">> " & sdData

#If SeguridadAlkon Then
    bK = CheckSum(bK, sdData)


    'Agregamos el fin de linea
    sdData = sdData & "~" & bK & ENDC
#Else
    
    sdData = sdData & ENDC
#End If

    'Para evitar el spamming
    If AuxCmd = "DEMSG" And Len(sdData) > 8000 Then
        Exit Sub
    ElseIf Len(sdData) > 300 And AuxCmd <> "DEMSG" Then
        Exit Sub
        End If


#If UsarWrench = 1 Then
    Call frmMain.Socket1.Write(sdData, Len(sdData))
#Else
    Call frmMain.Winsock1.SendData(sdData)

#End If

End Sub

Sub Login()
    If EstadoLogin = Normal Then
        SendData ("OOLOGI" & PJClickeado & "," & nombrecuent)
    ElseIf EstadoLogin = CrearNuevoPj Then
        SendData ("NLOGIN" & UserName & "," & UserRaza & "," & UserSexo & "," & UserSexo & "," & UserClase & "," & UserHogar _
                & "," & UserSkills(1) & "," & UserSkills(2) _
                & "," & UserSkills(3) & "," & UserSkills(4) _
                & "," & UserSkills(5) & "," & UserSkills(6) _
                & "," & UserSkills(7) & "," & UserSkills(8) _
                & "," & UserSkills(9) & "," & UserSkills(10) _
                & "," & UserSkills(11) & "," & UserSkills(12) _
                & "," & UserSkills(13) & "," & UserSkills(14) _
                & "," & UserSkills(15) & "," & UserSkills(16) _
                & "," & UserSkills(17) & "," & UserSkills(18) _
                & "," & UserSkills(19) & "," & UserSkills(20) _
                & "," & UserSkills(21) & "," & nombrecuent)
     ElseIf EstadoLogin = CrearAccount Then
     With frmCrearAccount
        SendData ("NACCNT" & .Nombre & "," & .Pass & "," & .Mail & "," & .pregunta & "," & .respuesta)
End With

    ElseIf EstadoLogin = BorrarPj Then
        SendData ("BORR" & PJClickeado)
    ElseIf EstadoLogin = LoginAccount Then
        SendData ("ALOGIN" & nombrecuent & "," & UserPassword & "," & App.Major & "." & App.Minor & "." & App.Revision & "," & MD5HushYo)
    End If
End Sub

