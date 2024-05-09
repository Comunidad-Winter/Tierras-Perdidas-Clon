Attribute VB_Name = "Mod_General"

Option Explicit

Public bK As Long
Public bRK As Long


Public iplst As String

Public bFogata As Boolean

Public bLluvia() As Byte ' Array para determinar si
'debemos mostrar la animacion de la lluvia

Public Function DirPath(ByVal path As String) As String
'•Parra: Nuevo Engine v2.0
    Select Case path
        Case "Graficos"
            DirPath = App.path & "\GRAFICOS\"
            Exit Function
        
        Case "Sound"
            DirPath = App.path & "\WAV\"
            Exit Function
        
        Case "Midi"
            DirPath = App.path & "\MIDI\"
            Exit Function
        
        Case "Maps"
            DirPath = App.path & "\MAPAS\"
            Exit Function
    End Select
End Function

Public Function DirGraficos() As String
    DirGraficos = App.path & "\" & "GRAFICOS" & "\"
End Function

Public Function DirSound() As String
    DirSound = App.path & "\" & "WAV" & "\"
End Function

Public Function DirMidi() As String
    DirMidi = App.path & "\" & "MIDI" & "\"
End Function

Public Function DirMapas() As String
    DirMapas = App.path & "\" & "MAPAS" & "\"
End Function

Public Function SumaDigitos(ByVal Numero As Integer) As Integer
    'Suma digitos
    Do
        SumaDigitos = SumaDigitos + (Numero Mod 10)
        Numero = Numero \ 10
    Loop While (Numero > 0)
End Function

Public Function SumaDigitosMenos(ByVal Numero As Integer) As Integer
    'Suma digitos, y resta el total de dígitos
    Do
        SumaDigitosMenos = SumaDigitosMenos + (Numero Mod 10) - 1
        Numero = Numero \ 10
    Loop While (Numero > 0)
End Function

Public Function Complex(ByVal Numero As Integer) As Integer
    If Numero Mod 2 <> 0 Then
        Complex = Numero * SumaDigitos(Numero)
    Else
        Complex = Numero * SumaDigitosMenos(Numero)
    End If
End Function

Public Function ValidarLoginMSG(ByVal Numero As Integer) As Integer
    Dim AuxInteger As Integer
    Dim AuxInteger2 As Integer
    
    AuxInteger = SumaDigitos(Numero)
    AuxInteger2 = SumaDigitosMenos(Numero)
    ValidarLoginMSG = Complex(AuxInteger + AuxInteger2)
End Function

Public Function RandomNumber(ByVal LowerBound As Long, ByVal UpperBound As Long) As Long
    'Initialize randomizer
    Randomize timer
    
    'Generate random number
    RandomNumber = (UpperBound - LowerBound) * Rnd + LowerBound
End Function

Sub CargarAnimArmas()
On Error Resume Next

    Dim LoopC As Long
    Dim arch As String
    
    arch = App.path & "\init\" & "armas.dat"
    
    NumWeaponAnims = Val(GetVar(arch, "INIT", "NumArmas"))
    
    ReDim WeaponAnimData(1 To NumWeaponAnims) As WeaponAnimData
    
    For LoopC = 1 To NumWeaponAnims
        InitGrh WeaponAnimData(LoopC).WeaponWalk(1), Val(GetVar(arch, "ARMA" & LoopC, "Dir1")), 0
        InitGrh WeaponAnimData(LoopC).WeaponWalk(2), Val(GetVar(arch, "ARMA" & LoopC, "Dir2")), 0
        InitGrh WeaponAnimData(LoopC).WeaponWalk(3), Val(GetVar(arch, "ARMA" & LoopC, "Dir3")), 0
        InitGrh WeaponAnimData(LoopC).WeaponWalk(4), Val(GetVar(arch, "ARMA" & LoopC, "Dir4")), 0
    Next LoopC
End Sub

Sub CargarVersiones()
On Error GoTo errorH:

    Versiones(1) = Val(GetVar(App.path & "\init\" & "versiones.ini", "Graficos", "Val"))
    Versiones(2) = Val(GetVar(App.path & "\init\" & "versiones.ini", "Wavs", "Val"))
    Versiones(3) = Val(GetVar(App.path & "\init\" & "versiones.ini", "Midis", "Val"))
    Versiones(4) = Val(GetVar(App.path & "\init\" & "versiones.ini", "Init", "Val"))
    Versiones(5) = Val(GetVar(App.path & "\init\" & "versiones.ini", "Mapas", "Val"))
    Versiones(6) = Val(GetVar(App.path & "\init\" & "versiones.ini", "E", "Val"))
    Versiones(7) = Val(GetVar(App.path & "\init\" & "versiones.ini", "O", "Val"))
Exit Sub

errorH:
    Call MsgBox("Error cargando versiones")
End Sub

Sub CargarColores()
    Dim archivoC As String
    
    archivoC = App.path & "\init\colores.dat"
    
    If Not FileExist(archivoC, vbArchive) Then
'TODO : Si hay que reinstalar, porque no cierra???
        Call MsgBox("ERROR: no se ha podido cargar los colores. Falta el archivo colores.dat, reinstale el juego", vbCritical + vbOKOnly)
        Exit Sub
    End If
    
    Dim i As Long
    
    For i = 0 To 48 '49 y 50 reservados para ciudadano y criminal
        ColoresPJ(i).r = CByte(GetVar(archivoC, CStr(i), "R"))
        ColoresPJ(i).g = CByte(GetVar(archivoC, CStr(i), "G"))
        ColoresPJ(i).B = CByte(GetVar(archivoC, CStr(i), "B"))
    Next i
    
    ColoresPJ(50).r = CByte(GetVar(archivoC, "CR", "R"))
    ColoresPJ(50).g = CByte(GetVar(archivoC, "CR", "G"))
    ColoresPJ(50).B = CByte(GetVar(archivoC, "CR", "B"))
    ColoresPJ(49).r = CByte(GetVar(archivoC, "CI", "R"))
    ColoresPJ(49).g = CByte(GetVar(archivoC, "CI", "G"))
    ColoresPJ(49).B = CByte(GetVar(archivoC, "CI", "B"))
End Sub


Sub CargarAnimEscudos()
On Error Resume Next

    Dim LoopC As Long
    Dim arch As String
    
    arch = App.path & "\init\" & "escudos.dat"
    
    NumEscudosAnims = Val(GetVar(arch, "INIT", "NumEscudos"))
    
    ReDim ShieldAnimData(1 To NumEscudosAnims) As ShieldAnimData
    
    For LoopC = 1 To NumEscudosAnims
        InitGrh ShieldAnimData(LoopC).ShieldWalk(1), Val(GetVar(arch, "ESC" & LoopC, "Dir1")), 0
        InitGrh ShieldAnimData(LoopC).ShieldWalk(2), Val(GetVar(arch, "ESC" & LoopC, "Dir2")), 0
        InitGrh ShieldAnimData(LoopC).ShieldWalk(3), Val(GetVar(arch, "ESC" & LoopC, "Dir3")), 0
        InitGrh ShieldAnimData(LoopC).ShieldWalk(4), Val(GetVar(arch, "ESC" & LoopC, "Dir4")), 0
    Next LoopC
End Sub

Sub AddtoRichTextBox(ByRef RichTextBox As RichTextBox, ByVal Text As String, Optional ByVal red As Integer = -1, Optional ByVal green As Integer, Optional ByVal blue As Integer, Optional ByVal bold As Boolean = False, Optional ByVal italic As Boolean = False, Optional ByVal bCrLf As Boolean = False)
'******************************************
'Adds text to a Richtext box at the bottom.
'Automatically scrolls to new text.
'Text box MUST be multiline and have a 3D
'apperance!
'******************************************
If Usuario.UserConsola = 0 Then
    
    With RichTextBox
        If (Len(.Text)) > 10000 Then .Text = ""
        
        .SelStart = Len(RichTextBox.Text)
        .SelLength = 0
        
        .SelBold = bold
        .SelItalic = italic
        
        If Not red = -1 Then .SelColor = RGB(red, green, blue)
        
        .SelText = IIf(bCrLf, Text, Text & vbCrLf)
        
        RichTextBox.Refresh
    End With
End If
End Sub

'TODO : Never was sure this is really necessary....
'TODO : 08/03/2006 - (AlejoLp) Esto hay que volarlo...
Public Sub RefreshAllChars()
'*****************************************************************
'Goes through the charlist and replots all the characters on the map
'Used to make sure everyone is visible
'*****************************************************************
    Dim LoopC As Long
    
    For LoopC = 1 To LastChar
        If charlist(LoopC).active = 1 Then
            MapData(charlist(LoopC).Pos.X, charlist(LoopC).Pos.Y).charindex = LoopC
        End If
    Next LoopC
End Sub

Sub SaveGameini()
    'Grabamos los datos del usuario en el Game.ini
    Config_Inicio.Name = "BetaTester"
    Config_Inicio.Password = "DammLamers"
    Config_Inicio.Puerto = UserPort
    
    Call EscribirGameIni(Config_Inicio)
End Sub

Function AsciiValidos(ByVal cad As String) As Boolean
    Dim car As Byte
    Dim i As Long
    
    cad = LCase$(cad)
    
    For i = 1 To Len(cad)
        car = Asc(mid$(cad, i, 1))
        
        If ((car < 97 Or car > 122) Or car = Asc("º")) And (car <> 255) And (car <> 32) Then
            Exit Function
        End If
    Next i
    
    AsciiValidos = True
End Function

Function CheckUserData(ByVal checkemail As Boolean) As Boolean
    'Validamos los datos del user
    Dim LoopC As Long
    Dim CharAscii As Integer
    
    If checkemail And UserEmail = "" Then
        frmMensaje.Show
        frmMensaje.MSG.Caption = ("Dirección de email invalida")
        Exit Function
    End If
    
    For LoopC = 1 To Len(UserPassword)
        CharAscii = Asc(mid$(UserPassword, LoopC, 1))
        If Not LegalCharacter(CharAscii) Then
            MsgBox ("Password inválido. El caractér " & Chr$(CharAscii) & " no está permitido.")
            Exit Function
        End If
    Next LoopC
    
    If nombrecuent = "" Then
        frmMensaje.Show
        frmMensaje.MSG.Caption = ("Ingrese un nombre de cuenta.")
        Exit Function
    End If
    
    If UserPassword = "" Then
        frmMensaje.Show
        frmMensaje.MSG.Caption = ("Ingrese un password.")
        Exit Function
    End If
    If Len(nombrecuent) > 30 Then
        frmMensaje.Show
        frmMensaje.MSG.Caption = ("La cuenta debe tener menos de 30 letras.")
        Exit Function
    End If
    
    For LoopC = 1 To Len(nombrecuent)
        CharAscii = Asc(mid$(nombrecuent, LoopC, 1))
        If Not LegalCharacter(CharAscii) Then
            frmMensaje.Show
        frmMensaje.MSG.Caption = ("Cuenta inválida. El caractér " & Chr$(CharAscii) & " no está permitido.")
            Exit Function
        End If
    Next LoopC
    
    CheckUserData = True
End Function

Sub UnloadAllForms()
On Error Resume Next

    Dim mifrm As Form

    For Each mifrm In Forms
        Unload mifrm

    Next
End Sub

Function LegalCharacter(ByVal KeyAscii As Integer) As Boolean
'*****************************************************************
'Only allow characters that are Win 95 filename compatible
'*****************************************************************
    'if backspace allow
    If KeyAscii = 8 Then
        LegalCharacter = True
        Exit Function
    End If
    
    'Only allow space, numbers, letters and special characters
    If KeyAscii < 32 Or KeyAscii = 44 Then
        Exit Function
    End If
    
    If KeyAscii > 126 Then
        Exit Function
    End If
    
    'Check for bad special characters in between
    If KeyAscii = 34 Or KeyAscii = 42 Or KeyAscii = 47 Or KeyAscii = 58 Or KeyAscii = 60 Or KeyAscii = 62 Or KeyAscii = 63 Or KeyAscii = 92 Or KeyAscii = 124 Then
        Exit Function
    End If
    
    'else everything is cool
    LegalCharacter = True
End Function

Sub SetConnected()
'*****************************************************************
'Sets the client to "Connect" mode
'*****************************************************************
    'Set Connected
    Connected = True
    
    Call SaveGameini
    'Unload the connect form
    Unload frmConnect
    
    
    'Load main form
    frmMain.Visible = True
frmMain.UserName.Caption = PJClickeado
Call DibujarPuntoMinimap
Call DibujarMinimap

If UserLvl > 50 Then
frmMain.LvlLbl.ForeColor = vbYellow
Else
frmMain.LvlLbl.ForeColor = vbRed
End If

Rem Cargamos Opciones Generales...
'If frmMain.Visible = True Then
    Dim Activado As Integer
    Activado = Val(GetVar(App.path & "\INIT\UserOptions.ini", "Opciones", "MenuPJs"))
    frmOpciones.Check1.value = Activado
    
    Activado = Val(GetVar(App.path & "\INIT\UserOptions.ini", "Opciones", "Sombras"))
    frmOpciones.Sombras.value = Activado
    
    Activado = Val(GetVar(App.path & "\INIT\UserOptions.ini", "Opciones", "Musica"))
    Musica = Activado
    
    Activado = Val(GetVar(App.path & "\INIT\UserOptions.ini", "Opciones", "Sonidos"))
    Sound = Activado
    
    Activado = Val(GetVar(App.path & "\INIT\UserOptions.ini", "Opciones", "LuzMouse"))
    frmOpciones.Check2.value = Activado
    If frmOpciones.Check2.value = 1 Then
    LuzMouse = True
    End If
    
    Activado = Val(GetVar(App.path & "\INIT\UserOptions.ini", "Opciones", "Minimap"))
    frmOpciones.Minimap.value = Activado
    
    Activado = Val(GetVar(App.path & "\INIT\UserOptions.ini", "Opciones", "Reflejo"))
    frmOpciones.optReflejo.value = Activado
    
    Activado = Val(GetVar(App.path & "\INIT\UserOptions.ini", "Opciones", "Alpha"))
    frmOpciones.OptTrans.value = Activado
    
If frmOpciones.Minimap.value = Unchecked Then
frmMain.Minimap.Visible = False
End If
If frmOpciones.Minimap.value = Checked Then
frmMain.Minimap.Visible = True
End If
    
    If frmOpciones.Check2.value = 0 Then
        Light.Light_Remove (Light.Light_Find(20))

End If
If frmOpciones.Check2.value = Checked Then
        Light.Light_Create UserPos.X + frmMain.MouseX \ 32 - frmMain.renderer.ScaleWidth \ 64, UserPos.Y + frmMain.MouseY / 32 - frmMain.renderer.ScaleHeight \ 64, 3, 20, 255, 255, 255
    End If
    
'End If

End Sub

Sub CargarTip()
    Dim N As Integer
    N = RandomNumber(1, UBound(Tips))
    
    frmtip.tip.Caption = Tips(N)
End Sub

Sub MoveTo(ByVal Direccion As E_Heading)
'***************************************************
'Author: Alejandro Santos (AlejoLp)
'Last Modify Date: 06/03/2006
' 06/03/2006: AlejoLp - Elimine las funciones Move[NSWE] y las converti a esta
'***************************************************
    Dim LegalOk As Boolean
    
    If Cartel Then Cartel = False
    
    Select Case Direccion
        Case E_Heading.NORTH
            LegalOk = LegalPos(UserPos.X, UserPos.Y - 1)
        Case E_Heading.EAST
            LegalOk = LegalPos(UserPos.X + 1, UserPos.Y)
        Case E_Heading.SOUTH
            LegalOk = LegalPos(UserPos.X, UserPos.Y + 1)
        Case E_Heading.WEST
            LegalOk = LegalPos(UserPos.X - 1, UserPos.Y)
    End Select
    
    If LegalOk Then
        Call SendData("M" & Direccion)
        DibujarPuntoMinimap
        If Not UserDescansar And Not UserMeditar And Not UserParalizado Then
            engine.Char_Move_by_Head UserCharIndex, Direccion
            MoveScreen Direccion
        End If
    Else
        If charlist(UserCharIndex).Heading <> Direccion Then
            Call SendData("CHEA" & Direccion)
        End If
    End If
    
End Sub

Sub RandomMove()
'***************************************************
'Author: Alejandro Santos (AlejoLp)
'Last Modify Date: 06/03/2006
' 06/03/2006: AlejoLp - Ahora utiliza la funcion MoveTo
'***************************************************

    MoveTo RandomNumber(1, 4)
    
End Sub

Sub CheckKeys()
'*****************************************************************
'Checks keys and respond
'*****************************************************************
On Error Resume Next
    'Don't allow any these keys during movement..
    If UserMoving = 0 Then
        If Not UserEstupido Then
            'Move Up
            If GetKeyState(vbKeyUp) < 0 Then
              
                Call MoveTo(NORTH)
                frmMain.Coord.Caption = MapName & " (" & UserMap & "," & UserPos.X & "," & UserPos.Y & ")"
                Exit Sub
            End If
        
            'Move Right
            If GetKeyState(vbKeyRight) < 0 Then
              
                Call MoveTo(EAST)
                frmMain.Coord.Caption = MapName & " (" & UserMap & "," & UserPos.X & "," & UserPos.Y & ")"
                Exit Sub
            End If
        
            'Move down
            If GetKeyState(vbKeyDown) < 0 Then
              
                Call MoveTo(SOUTH)
                frmMain.Coord.Caption = MapName & " (" & UserMap & "," & UserPos.X & "," & UserPos.Y & ")"
                Exit Sub
            End If
        
            'Move left
            If GetKeyState(vbKeyLeft) < 0 Then
             
                Call MoveTo(WEST)
               frmMain.Coord.Caption = MapName & " (" & UserMap & "," & UserPos.X & "," & UserPos.Y & ")"
                Exit Sub
            End If
        Else
            Dim kp As Boolean
            kp = (GetKeyState(vbKeyUp) < 0) Or _
                GetKeyState(vbKeyRight) < 0 Or _
                GetKeyState(vbKeyDown) < 0 Or _
                GetKeyState(vbKeyLeft) < 0
            If kp Then Call RandomMove
          
            frmMain.Coord.Caption = MapName & " (" & UserMap & "," & UserPos.X & "," & UserPos.Y & ")"
        End If
    End If
End Sub

'TODO : esto no es del tileengine??
Sub MoveScreen(ByVal nHeading As E_Heading)
'******************************************
'Starts the screen moving in a direction
'******************************************
    Dim X As Integer
    Dim Y As Integer
    Dim tX As Integer
    Dim tY As Integer
    
    'Figure out which way to move
    Select Case nHeading
        Case E_Heading.NORTH
            Y = -1
    
        Case E_Heading.EAST
            X = 1
    
        Case E_Heading.SOUTH
            Y = 1
        
        Case E_Heading.WEST
            X = -1
            
    End Select
    
    'Fill temp pos
    tX = UserPos.X + X
    tY = UserPos.Y + Y

    If Not (tX < MinXBorder Or tX > MaxXBorder Or tY < MinYBorder Or tY > MaxYBorder) Then
        AddtoUserPos.X = X
        UserPos.X = tX
        AddtoUserPos.Y = Y
        UserPos.Y = tY
        UserMoving = 1
        
        bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or _
                MapData(UserPos.X, UserPos.Y).Trigger = 2 Or _
                MapData(UserPos.X, UserPos.Y).Trigger = 4, True, False)
        Exit Sub
    End If
End Sub

'TODO : esto no es del tileengine??
Function NextOpenChar()
'******************************************
'Finds next open Char
'******************************************
    Dim LoopC As Long
    
    LoopC = 1
    Do While charlist(LoopC).active And LoopC < UBound(charlist)
        LoopC = LoopC + 1
    Loop
    
    NextOpenChar = LoopC
End Function

Sub SwitchMap(ByVal Map As Integer)
'**************************************************************
'Formato de mapas optimizado para reducir el espacio que ocupan.
'Diseñado y creado por Juan Martín Sotuyo Dodero (Maraxus) (juansotuyo@hotmail.com)
'**************************************************************
    Dim Y As Long
    Dim X As Long
    Dim tempint As Integer
    Dim ByFlags As Byte
    Dim handle As Integer
    Dim TempLng As Byte
    Dim TempByte1 As Byte
    Dim TempByte2 As Byte
    Dim TempByte3 As Byte
    Dim i As Byte

    'By Lorwik - www.rincondelao.com.ar
    engine.Particle_Group_Remove_All
    Light.Light_Remove_All
    handle = FreeFile()
    
    Open DirPath("Maps") & "Mapa" & Map & ".map" For Binary As handle
    Seek handle, 1
            
    'map Header
    Get handle, , MapInfo.MapVersion
    Get handle, , MiCabecera
    Get handle, , tempint
    Get handle, , tempint
    Get handle, , tempint
    Get handle, , tempint
    
    'Load arrays
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            For i = 0 To 3
                'By Lorwik - www.rincondelao.com.ar
                MapData(X, Y).light_value(i) = False
            Next i
            Get handle, , ByFlags
            MapData(X, Y).luz = 0
            MapData(X, Y).particle_group = 0
            MapData(X, Y).Blocked = (ByFlags And 1)
            
            Get handle, , MapData(X, Y).Graphic(1).grhindex
            InitGrh MapData(X, Y).Graphic(1), MapData(X, Y).Graphic(1).grhindex
            
            'Layer 2 used?
            If ByFlags And 2 Then
                Get handle, , MapData(X, Y).Graphic(2).grhindex
                InitGrh MapData(X, Y).Graphic(2), MapData(X, Y).Graphic(2).grhindex
            Else
                MapData(X, Y).Graphic(2).grhindex = 0
            End If
                
            'Layer 3 used?
            If ByFlags And 4 Then
                Get handle, , MapData(X, Y).Graphic(3).grhindex
                InitGrh MapData(X, Y).Graphic(3), MapData(X, Y).Graphic(3).grhindex
            Else
                MapData(X, Y).Graphic(3).grhindex = 0
            End If
                
            'Layer 4 used?
            If ByFlags And 8 Then
                Get handle, , MapData(X, Y).Graphic(4).grhindex
                InitGrh MapData(X, Y).Graphic(4), MapData(X, Y).Graphic(4).grhindex
            Else
                MapData(X, Y).Graphic(4).grhindex = 0
            End If
            
            'Trigger used?
            If ByFlags And 16 Then
                Get handle, , MapData(X, Y).Trigger
            Else
                MapData(X, Y).Trigger = 0
            End If
                        If ByFlags And 32 Then
               Get handle, , tempint
                'By Lorwik - www.rincondelao.com.ar
                MapData(X, Y).particle_group_index = engine.General_Particle_Create(tempint, X, Y, -1)
            End If
            
            If ByFlags And 64 Then
                'By Lorwik - www.rincondelao.com.ar
                Get handle, , MapData(X, Y).base_light(0)
                Get handle, , MapData(X, Y).base_light(1)
                Get handle, , MapData(X, Y).base_light(2)
                Get handle, , MapData(X, Y).base_light(3)
                
                If MapData(X, Y).base_light(0) Then _
                    Get handle, , MapData(X, Y).light_value(0)
                
                If MapData(X, Y).base_light(1) Then _
                    Get handle, , MapData(X, Y).light_value(1)
                 'By Lorwik - www.rincondelao.com.ar
                If MapData(X, Y).base_light(2) Then _
                    Get handle, , MapData(X, Y).light_value(2)
                
                If MapData(X, Y).base_light(3) Then _
                    Get handle, , MapData(X, Y).light_value(3)
            End If
            'Erase NPCs
            If MapData(X, Y).charindex > 0 Then
                Call EraseChar(MapData(X, Y).charindex)
            End If
            
            'Erase OBJs
            MapData(X, Y).ObjGrh.grhindex = 0
        Next X
    Next Y
    
    Close handle
    
    MapInfo.Name = ""
    MapInfo.Music = ""
    
    'CurMap = Map
   If frmConnect.Visible Then Exit Sub
        Call DibujarPuntoMinimap
        Call DibujarMinimap
End Sub

'TODO : Reemplazar por la nueva versión, esta apesta!!!
Public Function ReadField(ByVal Pos As Integer, ByVal Text As String, ByVal SepASCII As Integer) As String
'*****************************************************************
'Gets a field from a string
'*****************************************************************
    Dim i As Integer
    Dim LastPos As Integer
    Dim CurChar As String * 1
    Dim FieldNum As Integer
    Dim Seperator As String
    
    Seperator = Chr$(SepASCII)
    LastPos = 0
    FieldNum = 0
    
    For i = 1 To Len(Text)
        CurChar = mid$(Text, i, 1)
        If CurChar = Seperator Then
            FieldNum = FieldNum + 1
            If FieldNum = Pos Then
                ReadField = mid$(Text, LastPos + 1, (InStr(LastPos + 1, Text, Seperator, vbTextCompare) - 1) - (LastPos))
                Exit Function
            End If
            LastPos = i
        End If
    Next i
    FieldNum = FieldNum + 1
    
    If FieldNum = Pos Then
        ReadField = mid$(Text, LastPos + 1)
    End If
End Function

Function FileExist(ByVal file As String, ByVal FileType As VbFileAttribute) As Boolean
    FileExist = (Dir$(file, FileType) <> "")
End Function

Sub WriteClientVer()
    Dim hFile As Integer
        
    hFile = FreeFile()
    Open App.path & "\init\Ver.bin" For Binary Access Write Lock Read As #hFile
    Put #hFile, , CLng(777)
    Put #hFile, , CLng(777)
    Put #hFile, , CLng(777)
    
    Put #hFile, , CInt(App.Major)
    Put #hFile, , CInt(App.Minor)
    Put #hFile, , CInt(App.Revision)
    
    Close #hFile
End Sub

Public Function IsIp(ByVal Ip As String) As Boolean
    Dim i As Long
    
    For i = 1 To UBound(ServersLst)
        If ServersLst(i).Ip = Ip Then
            IsIp = True
            Exit Function
        End If
    Next i
End Function

Public Sub CargarServidores()
On Error GoTo errorH
    Dim f As String
    Dim c As Integer
    Dim i As Long
    
    f = App.path & "\init\sinfo.dat"
    c = Val(GetVar(f, "INIT", "Cant"))
    
    ReDim ServersLst(1 To c) As tServerInfo
    For i = 1 To c
        ServersLst(i).desc = GetVar(f, "S" & i, "Desc")
        ServersLst(i).Ip = Trim$(GetVar(f, "S" & i, "Ip"))
        ServersLst(i).PassRecPort = CInt(GetVar(f, "S" & i, "P2"))
        ServersLst(i).Puerto = CInt(GetVar(f, "S" & i, "PJ"))
    Next i
    CurServer = 1
Exit Sub

errorH:
    Call MsgBox("Error cargando los servidores, actualicelos de la web", vbCritical + vbOKOnly, "Argentum Online")
    End
End Sub

Public Sub InitServersList(ByVal Lst As String)
On Error Resume Next
    Dim NumServers As Integer
    Dim i As Integer
    Dim Cont As Integer
    
    i = 1
    
    Do While (ReadField(i, RawServersList, Asc(";")) <> "")
        i = i + 1
        Cont = Cont + 1
    Loop
    
    ReDim ServersLst(1 To Cont) As tServerInfo
    
    For i = 1 To Cont
        Dim cur$
        cur$ = ReadField(i, RawServersList, Asc(";"))
        ServersLst(i).Ip = ReadField(1, cur$, Asc(":"))
        ServersLst(i).Puerto = ReadField(2, cur$, Asc(":"))
        ServersLst(i).desc = ReadField(4, cur$, Asc(":"))
        ServersLst(i).PassRecPort = ReadField(3, cur$, Asc(":"))
    Next i
    
    CurServer = 1
End Sub

Public Function CurServerPasRecPort() As Integer

        CurServerPasRecPort = "7667"

End Function

Public Function CurServerIp() As String

        CurServerIp = "127.0.0.1"

End Function

Public Function CurServerPort() As Integer

        CurServerPort = "7666"

End Function


Sub Main()

On Error Resume Next
Set Light = New clsLight
    Call WriteClientVer

    If App.PrevInstance Then
        Call MsgBox("Argentum Online ya esta corriendo! No es posible correr otra instancia del juego. Haga click en Aceptar para salir.", vbApplicationModal + vbInformation + vbOKOnly, "Error al ejecutar")
        End
    End If

    
    ChDrive App.path
    ChDir App.path

    MD5HushYo = "0123456789abcdef"  'We aren't using a real MD5
    
    'Por default usamos el dinámico
    Set SurfaceDB = New clsSurfaceManDynDX8

        
    frmCargando.Show
    Call frmCargando.establecerProgreso(0)
    frmCargando.LabelCarga.Caption = "¡¡¡Bienvenido a Lhirius AO!!!"
If MsgBox("¿Deseas jugar en modo ventana?", vbYesNo, "Resolucion") = vbNo Then
Call Resolucion.SetResolucion
End If
    frmCargando.Refresh
    
    frmMain.Socket1.Startup
    Call InicializarNombres
    Call frmCargando.progresoConDelay(10)
    UserMap = 1
    frmCargando.LabelCarga.Caption = "Cargando Indices..."
    LoadGrhData
    CargarCabezas
    Call frmCargando.progresoConDelay(20)
    CargarCascos
    CargarCuerpos
    Call frmCargando.progresoConDelay(30)
    CargarArrayLluvia
    Call frmCargando.progresoConDelay(40)
    CargarFxs
    Call frmCargando.progresoConDelay(50)
    Call engine.Engine_Init
    frmCargando.LabelCarga.Caption = "Cargando Engine Grafico..."
    Call frmCargando.progresoConDelay(60)
    Call engine.setup_ambient
    frmCargando.LabelCarga.Caption = "Cargando Particulas..."
    Call CargarParticulas
    Call frmCargando.progresoConDelay(70)
    Call CargarArrayLluvia
    frmCargando.LabelCarga.Caption = "Cargando Animaciones..."
    Call CargarAnimArmas
    Call frmCargando.progresoConDelay(80)
    Call CargarAnimEscudos
    Call CargarVersiones
    Call frmCargando.progresoConDelay(90)
    Call CargarColores
    Call CargarFonts
    Call frmCargando.progresoConDelay(95)
    frmCargando.LabelCarga.Caption = "Buscando Actualizaciones..."
    frmCargando.Analizar
    Call frmCargando.progresoConDelay(96)
    frmCargando.LabelCarga.Caption = "¡¡¡Bienvenido a Lhirius AO!!!"
    Call frmCargando.progresoConDelay(100)
    
    Unload frmCargando
            
    'Inicializamos el sonido
     Call Audio.Initialize(frmMain.hWnd, App.path & "\WAV\", App.path & "\MIDI\")
    
    'Inicializamos el inventario gráfico
    Call Inventario.Initialize(frmMain.picInv)
    
        Call Audio.PlayWave("313.wav")


    'frmPres.Picture = LoadPicture(App.path & "\Graficos\pres" & RandomNumber(1, 5) & ".jpg")
    frmPres.Show vbModal    'Es modal, así que se detiene la ejecución de Main hasta que se desaparece
    
    frmConnect.Visible = True

    'Inicialización de variables globales
    prgRun = True
    pausa = False
    
        Dialogos.font = frmMain.font
    DialogosClanes.font = frmMain.font
engine.Start

Exit Sub
ManejadorErrores:
    MsgBox "Ha ocurrido un error irreparable, el cliente se cerrará."
    Debug.Print "Contexto:" & Err.HelpContext & " Desc:" & Err.Description & " Fuente:" & Err.Source
    End
End Sub
Function FieldCount(ByRef Text As String, ByVal SepASCII As Byte) As Long
'*****************************************************************
'Gets the number of fields in a delimited string
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 07/29/2007
'*****************************************************************
    Dim Count As Long
    Dim curPos As Long
    Dim delimiter As String * 1
    
    If LenB(Text) = 0 Then Exit Function
    
    delimiter = Chr$(SepASCII)
    
    curPos = 0
    
    Do
        curPos = InStr(curPos + 1, Text, delimiter)
        Count = Count + 1
    Loop While curPos <> 0
    
    FieldCount = Count
End Function

Sub WriteVar(ByVal file As String, ByVal Main As String, ByVal var As String, ByVal value As String)
'*****************************************************************
'Writes a var to a text file
'*****************************************************************
    writeprivateprofilestring Main, var, value, file
End Sub

Function GetVar(ByVal file As String, ByVal Main As String, ByVal var As String) As String
'*****************************************************************
'Gets a Var from a text file
'*****************************************************************
    Dim sSpaces As String ' This will hold the input that the program will retrieve
    
    sSpaces = Space$(100) ' This tells the computer how long the longest string can be. If you want, you can change the number 100 to any number you wish
    
    getprivateprofilestring Main, var, vbNullString, sSpaces, Len(sSpaces), file
    
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function

'[CODE 002]:MatuX
'
'  Función para chequear el email
'
'  Corregida por Maraxus para que reconozca como válidas casillas con puntos antes de la arroba y evitar un chequeo innecesario
Public Function CheckMailString(ByVal sString As String) As Boolean
On Error GoTo errHnd
    Dim lPos  As Long
    Dim lX    As Long
    Dim iAsc  As Integer
    
    '1er test: Busca un simbolo @
    lPos = InStr(sString, "@")
    If (lPos <> 0) Then
        '2do test: Busca un simbolo . después de @ + 1
        If Not (InStr(lPos, sString, ".", vbBinaryCompare) > lPos + 1) Then _
            Exit Function
        
        '3er test: Recorre todos los caracteres y los valída
        For lX = 0 To Len(sString) - 1
            If Not (lX = (lPos - 1)) Then   'No chequeamos la '@'
                iAsc = Asc(mid$(sString, (lX + 1), 1))
                If Not CMSValidateChar_(iAsc) Then _
                    Exit Function
            End If
        Next lX
        
        'Finale
        CheckMailString = True
    End If
errHnd:
End Function

'  Corregida por Maraxus para que reconozca como válidas casillas con puntos antes de la arroba
Private Function CMSValidateChar_(ByVal iAsc As Integer) As Boolean
    CMSValidateChar_ = (iAsc >= 48 And iAsc <= 57) Or _
                        (iAsc >= 65 And iAsc <= 90) Or _
                        (iAsc >= 97 And iAsc <= 122) Or _
                        (iAsc = 95) Or (iAsc = 45) Or (iAsc = 46)
End Function

'TODO : como todo lorelativo a mapas, no tiene anda que hacer acá....
Function HayAgua(ByVal X As Integer, ByVal Y As Integer) As Boolean
On Error Resume Next
    HayAgua = MapData(X, Y).Graphic(1).grhindex >= 1505 And _
                MapData(X, Y).Graphic(1).grhindex <= 1520 And _
                MapData(X, Y).Graphic(2).grhindex = 0
End Function

Public Sub ShowSendTxt()
    If Not frmCantidad.Visible Then
        frmMain.SendTxt.Visible = True
        frmMain.SendTxt.SetFocus
    End If
End Sub

Public Sub ShowSendCMSGTxt()
    If Not frmCantidad.Visible Then
        frmMain.SendCMSTXT.Visible = True
        frmMain.SendCMSTXT.SetFocus
    End If
End Sub
    
Public Sub LeerLineaComandos()
    Dim T() As String
    Dim i As Long
    
    'Parseo los comandos
    T = Split(Command, " ")
    
    For i = LBound(T) To UBound(T)
        Select Case UCase$(T(i))
            Case "/NORES" 'no cambiar la resolucion
                NoRes = True
        End Select
    Next i
End Sub

Private Sub LoadClientSetup()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/27/2005
'
'**************************************************************
    Dim fHandle As Integer
    
    fHandle = FreeFile
    Open App.path & "\init\ao.dat" For Binary Access Read Lock Write As fHandle
        Get fHandle, , ClientSetup
    Close fHandle
    
    Musica = Not ClientSetup.bNoMusic
    Sound = Not ClientSetup.bNoSound
End Sub

Private Sub InicializarNombres()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/27/2005
'Inicializa los nombres de razas, ciudades, clases, skills, atributos, etc.
'**************************************************************
    Ciudades(1) = "Ullathorpe"

    CityDesc(1) = "Ullathorpe está establecida en el medio de los grandes bosques de Argentum, es principalmente un pueblo de campesinos y leñadores. Su ubicación hace de Ullathorpe un punto de paso obligado para todos los aventureros ya que se encuentra cerca de los lugares más legendarios de este mundo."
    
    ListaRazas(1) = "Humano"
    ListaRazas(2) = "Elfo"
    ListaRazas(3) = "Elfo Oscuro"
    ListaRazas(4) = "Gnomo"
    ListaRazas(5) = "Enano"

    ListaClases(1) = "Mago"
    ListaClases(2) = "Clerigo"
    ListaClases(3) = "Guerrero"
    ListaClases(4) = "Asesino"
    ListaClases(5) = "Ladron"
    ListaClases(6) = "Bardo"
    ListaClases(7) = "Druida"
    ListaClases(8) = "Bandido"
    ListaClases(9) = "Paladin"
    ListaClases(10) = "Cazador"
    ListaClases(11) = "Pescador"
    ListaClases(12) = "Herrero"
    ListaClases(13) = "Leñador"
    ListaClases(14) = "Minero"
    ListaClases(15) = "Carpintero"
    ListaClases(16) = "Pirata"

    SkillsNames(Skills.Suerte) = "Suerte"
    SkillsNames(Skills.Magia) = "Magia"
    SkillsNames(Skills.Robar) = "Robar"
    SkillsNames(Skills.Tacticas) = "Tacticas de combate"
    SkillsNames(Skills.Armas) = "Combate con armas"
    SkillsNames(Skills.Meditar) = "Meditar"
    SkillsNames(Skills.Apuñalar) = "Apuñalar"
    SkillsNames(Skills.Ocultarse) = "Ocultarse"
    SkillsNames(Skills.Supervivencia) = "Supervivencia"
    SkillsNames(Skills.Talar) = "Talar árboles"
    SkillsNames(Skills.Comerciar) = "Comercio"
    SkillsNames(Skills.Defensa) = "Defensa con escudos"
    SkillsNames(Skills.Pesca) = "Pesca"
    SkillsNames(Skills.Mineria) = "Mineria"
    SkillsNames(Skills.Carpinteria) = "Carpinteria"
    SkillsNames(Skills.Herreria) = "Herreria"
    SkillsNames(Skills.Liderazgo) = "Liderazgo"
    SkillsNames(Skills.Domar) = "Domar animales"
    SkillsNames(Skills.Proyectiles) = "Armas de proyectiles"
    SkillsNames(Skills.Wresterling) = "Wresterling"
    SkillsNames(Skills.Navegacion) = "Navegacion"
    SkillsNames(Skills.Resistencia) = "Resistencia Magica"

    AtributosNames(1) = "Fuerza"
    AtributosNames(2) = "Agilidad"
    AtributosNames(3) = "Inteligencia"
    AtributosNames(4) = "Carisma"
    AtributosNames(5) = "Constitucion"
End Sub

Public Sub LogError(desc As String)
On Error Resume Next
Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.path & "\errores.log" For Append As #nfile
Print #nfile, desc
Close #nfile
End Sub

Public Sub LogCustom(desc As String)
On Error Resume Next
Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.path & "\custom.log" For Append As #nfile
Print #nfile, Now & " " & desc
Close #nfile
End Sub
Public Function PonerPuntos(Numero As Long) As String
Dim i As Integer
Dim Cifra As String
 
Cifra = str(Numero)
Cifra = Right$(Cifra, Len(Cifra) - 1)
For i = 0 To 4
    If Len(Cifra) - 3 * i >= 3 Then
        If mid$(Cifra, Len(Cifra) - (2 + 3 * i), 3) <> "" Then
            PonerPuntos = mid$(Cifra, Len(Cifra) - (2 + 3 * i), 3) & "." & PonerPuntos
        End If
    Else
        If Len(Cifra) - 3 * i > 0 Then
            PonerPuntos = Left$(Cifra, Len(Cifra) - 3 * i) & "." & PonerPuntos
        End If
        Exit For
    End If
Next
 
PonerPuntos = Left$(PonerPuntos, Len(PonerPuntos) - 1)
 
End Function
 
Public Function General_Var_Get(ByVal file As String, ByVal Main As String, ByVal var As String) As String
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Get a var to from a text file
'*****************************************************************
    Dim l As Long
    Dim Char As String
    Dim sSpaces As String 'Input that the program will retrieve
    Dim szReturn As String 'Default value if the string is not found
   
    szReturn = ""
   
    sSpaces = Space$(5000)
   
    getprivateprofilestring Main, var, szReturn, sSpaces, Len(sSpaces), file
   
    General_Var_Get = RTrim$(sSpaces)
    General_Var_Get = Left$(General_Var_Get, Len(General_Var_Get) - 1)
End Function
Public Function General_Field_Read(ByVal field_pos As Long, ByVal Text As String, ByVal delimiter As Byte) As String
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Gets a field from a delimited string
'*****************************************************************
    Dim i As Long
    Dim LastPos As Long
    Dim FieldNum As Long
   
    LastPos = 0
    FieldNum = 0
    For i = 1 To Len(Text)
        If delimiter = CByte(Asc(mid$(Text, i, 1))) Then
            FieldNum = FieldNum + 1
            If FieldNum = field_pos Then
                General_Field_Read = mid$(Text, LastPos + 1, (InStr(LastPos + 1, Text, Chr$(delimiter), vbTextCompare) - 1) - (LastPos))
                Exit Function
            End If
            LastPos = i
        End If
    Next i
    FieldNum = FieldNum + 1
    If FieldNum = field_pos Then
        General_Field_Read = mid$(Text, LastPos + 1)
    End If
End Function
Public Sub EquiparItem()
    If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then
       
        SendData "EQUI" & Inventario.SelectedItem
        End If
End Sub
Sub CargarParticulas()
Dim StreamFile As String
Dim LoopC As Long
Dim i As Long
Dim GrhListing As String
Dim TempSet As String
Dim ColorSet As Long
    
StreamFile = App.path & "\INIT\Particles.ini"
TotalStreams = Val(General_Var_Get(StreamFile, "INIT", "Total"))

'resize StreamData array
ReDim StreamData(1 To TotalStreams) As Stream

    'fill StreamData array with info from Particles.ini
    For LoopC = 1 To TotalStreams
        StreamData(LoopC).Name = General_Var_Get(StreamFile, Val(LoopC), "Name")
        StreamData(LoopC).NumOfParticles = General_Var_Get(StreamFile, Val(LoopC), "NumOfParticles")
        StreamData(LoopC).X1 = General_Var_Get(StreamFile, Val(LoopC), "X1")
        StreamData(LoopC).Y1 = General_Var_Get(StreamFile, Val(LoopC), "Y1")
        StreamData(LoopC).X2 = General_Var_Get(StreamFile, Val(LoopC), "X2")
        StreamData(LoopC).Y2 = General_Var_Get(StreamFile, Val(LoopC), "Y2")
        StreamData(LoopC).angle = General_Var_Get(StreamFile, Val(LoopC), "Angle")
        StreamData(LoopC).vecx1 = General_Var_Get(StreamFile, Val(LoopC), "VecX1")
        StreamData(LoopC).vecx2 = General_Var_Get(StreamFile, Val(LoopC), "VecX2")
        StreamData(LoopC).vecy1 = General_Var_Get(StreamFile, Val(LoopC), "VecY1")
        StreamData(LoopC).vecy2 = General_Var_Get(StreamFile, Val(LoopC), "VecY2")
        StreamData(LoopC).life1 = General_Var_Get(StreamFile, Val(LoopC), "Life1")
        StreamData(LoopC).life2 = General_Var_Get(StreamFile, Val(LoopC), "Life2")
        StreamData(LoopC).friction = General_Var_Get(StreamFile, Val(LoopC), "Friction")
        StreamData(LoopC).spin = General_Var_Get(StreamFile, Val(LoopC), "Spin")
        StreamData(LoopC).spin_speedL = General_Var_Get(StreamFile, Val(LoopC), "Spin_SpeedL")
        StreamData(LoopC).spin_speedH = General_Var_Get(StreamFile, Val(LoopC), "Spin_SpeedH")
        StreamData(LoopC).AlphaBlend = General_Var_Get(StreamFile, Val(LoopC), "AlphaBlend")
        StreamData(LoopC).gravity = General_Var_Get(StreamFile, Val(LoopC), "Gravity")
        StreamData(LoopC).grav_strength = General_Var_Get(StreamFile, Val(LoopC), "Grav_Strength")
        StreamData(LoopC).bounce_strength = General_Var_Get(StreamFile, Val(LoopC), "Bounce_Strength")
        StreamData(LoopC).XMove = General_Var_Get(StreamFile, Val(LoopC), "XMove")
        StreamData(LoopC).YMove = General_Var_Get(StreamFile, Val(LoopC), "YMove")
        StreamData(LoopC).move_x1 = General_Var_Get(StreamFile, Val(LoopC), "move_x1")
        StreamData(LoopC).move_x2 = General_Var_Get(StreamFile, Val(LoopC), "move_x2")
        StreamData(LoopC).move_y1 = General_Var_Get(StreamFile, Val(LoopC), "move_y1")
        StreamData(LoopC).move_y2 = General_Var_Get(StreamFile, Val(LoopC), "move_y2")
        StreamData(LoopC).life_counter = General_Var_Get(StreamFile, Val(LoopC), "life_counter")
        StreamData(LoopC).Speed = Val(General_Var_Get(StreamFile, Val(LoopC), "Speed"))
                StreamData(LoopC).grh_resize = Val(General_Var_Get(StreamFile, Val(LoopC), "resize"))
        StreamData(LoopC).grh_resizex = Val(General_Var_Get(StreamFile, Val(LoopC), "rx"))
        StreamData(LoopC).grh_resizey = Val(General_Var_Get(StreamFile, Val(LoopC), "ry"))
        StreamData(LoopC).NumGrhs = General_Var_Get(StreamFile, Val(LoopC), "NumGrhs")
        
        ReDim StreamData(LoopC).grh_list(1 To StreamData(LoopC).NumGrhs)
        GrhListing = General_Var_Get(StreamFile, Val(LoopC), "Grh_List")
        
        For i = 1 To StreamData(LoopC).NumGrhs
            StreamData(LoopC).grh_list(i) = General_Field_Read(str(i), GrhListing, 44)
        Next i
        StreamData(LoopC).grh_list(i - 1) = StreamData(LoopC).grh_list(i - 1)
        For ColorSet = 1 To 4
            TempSet = General_Var_Get(StreamFile, Val(LoopC), "ColorSet" & ColorSet)
            StreamData(LoopC).colortint(ColorSet - 1).r = General_Field_Read(1, TempSet, 44)
            StreamData(LoopC).colortint(ColorSet - 1).g = General_Field_Read(2, TempSet, 44)
            StreamData(LoopC).colortint(ColorSet - 1).B = General_Field_Read(3, TempSet, 44)
        Next ColorSet
            'frmMain.List2.AddItem loopc & " - " & StreamData(loopc).Name
    Next LoopC

End Sub



Public Sub DibujarPuntoMinimap()
    
With frmMain
.Puntito.Left = UserPos.X - 2
.Puntito.Top = UserPos.Y - 3
End With
    
End Sub
     
Public Sub DibujarMinimap()
If UserMap < 10 Then
frmMain.Minimap.Picture = LoadPicture(App.path & "\Graficos\2100" & UserMap & ".bmp")
End If
If UserMap < 100 Then
frmMain.Minimap.Picture = LoadPicture(App.path & "\Graficos\210" & UserMap & ".bmp")
End If
If UserMap >= 100 Then
frmMain.Minimap.Picture = LoadPicture(App.path & "\Graficos\21" & UserMap & ".bmp")
End If
End Sub
Public Function OpenBrowser(strURL As String, lngHwnd As Long)
    OpenBrowser = ShellExecute(lngHwnd, vbNullString, strURL, vbNullString, _
    "c:\", 0)
End Function
