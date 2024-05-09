Attribute VB_Name = "Mod_TileEngine"
'Argentum Online 0.9.0.9
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
'    C       O       N       S      T
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
'Map sizes in tiles
Public Const XMaxMapSize = 100

Public Const XMinMapSize = 1

Public Const YMaxMapSize = 100

Public Const YMinMapSize = 1

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
'    T       I      P      O      S
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

'Posicion en un mapa
Public Type Position

    x As Integer
    y As Integer

End Type

'Contiene info acerca de donde se puede encontrar un grh
'tamaño y animacion
Public Type GrhData

    sX As Integer
    sY As Integer
   
    FileNum As Long
   
    pixelWidth As Integer
    pixelHeight As Integer
   
    TileWidth As Single
    TileHeight As Single
   
    NumFrames As Integer
    Frames() As Long
   
    Speed As Single

End Type

'apunta a una estructura grhdata y mantiene la animacion
Public Type Grh

    grhindex As Integer
    FrameCounter As Single
    Speed As Single
    Started As Byte
    Loops As Integer

End Type

'Lista de cuerpos
Public Type BodyData

    Walk(1 To 4) As Grh
    HeadOffset As Position

End Type

'Lista de cabezas
Public Type HeadData

    Head(1 To 4) As Grh

End Type

'Lista de las animaciones de las armas
Type WeaponAnimData

    WeaponWalk(1 To 4) As Grh
    '[ANIM ATAK]
    WeaponAttack As Byte

End Type

'Lista de las animaciones de los escudos
Type ShieldAnimData

    ShieldWalk(1 To 4) As Grh

End Type

'Apariencia del personaje
Public Type Char
    particle_count As Integer
    particle_group() As Long
    TieneParticulas As Boolean
    FxIndex As Integer
    scrollDirectionX As Integer
    scrollDirectionY As Integer
   
    MoveOffsetX As Single
    MoveOffsetY As Single
    active As Byte
    Heading As E_Heading ' As E_Heading ?
    Pos As Position
    
    iHead As Integer
    iBody As Integer
    Body As BodyData
    Head As HeadData
    Casco As HeadData
    Arma As WeaponAnimData
    Escudo As ShieldAnimData
    UsandoArma As Boolean
    fX As Grh
    'FxLoopTimes As Integer
    Criminal As Byte
    
    Nombre As String
    
    Moving As Byte
    MoveOffset As Position
    'ServerIndex As Integer
    
    pie As Boolean
    Muerto As Boolean
    invisible As Boolean
    priv As Byte
    
End Type

'Tipo de las celdas del mapa
Public Type MapBlock
    light_value(3) As Long
    base_light(0 To 3) As Boolean 'Indica si el tile tiene luz propia.
    'particle_group As Integer
    particle_group_index As Integer

    luz As Integer
    'color(3) As Long
   
    particle_group As Integer
    Graphic(1 To 4) As Grh
    charindex As Integer
    ObjGrh As Grh
    
    'NPCIndex As Integer
    'OBJInfo As Obj
    'TileExit As WorldPos
    Blocked As Byte
    
    Trigger As Integer

End Type

'Info de cada mapa
Public Type MapInfo

    Music As String
    Name As String
    'StartPos As WorldPos
    MapVersion As Integer
    
    'ME Only
    'Changed As Byte

End Type

'Bordes del mapa
Public MinXBorder           As Byte

Public MaxXBorder           As Byte

Public MinYBorder           As Byte

Public MaxYBorder           As Byte

Public UserMoving           As Byte

Public UserPos              As Position 'Posicion

Public AddtoUserPos         As Position 'Si se mueve

Public UserCharIndex        As Integer

Public UserMaxAGU           As Integer

Public UserMinAGU           As Integer

Public UserMaxHAM           As Integer

Public UserMinHAM           As Integer

'Cuantos tiles el engine mete en el BUFFER cuando
'dibuja el mapa. Ojo un tamaño muy grande puede
'volver el engine muy lento
Public TileBufferSize       As Integer

'Tamaño de los tiles en pixels
Public TilePixelHeight      As Integer

Public TilePixelWidth       As Integer

'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Totales?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

Public NumChars             As Integer

Public LastChar             As Integer

Public NumWeaponAnims       As Integer

Public GrhData()            As GrhData 'Guarda todos los grh

Public BodyData()           As BodyData

Public HeadData()           As HeadData

Public FxData()             As tIndiceFx

Public WeaponAnimData()     As WeaponAnimData

Public ShieldAnimData()     As ShieldAnimData

Public CascoAnimData()      As HeadData
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Mapa?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
Public MapData()            As MapBlock ' Mapa

Public MapInfo              As MapInfo ' Info acerca del mapa en uso
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Usuarios?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
'
'epa ;)
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿API?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
'Blt
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
'       [CODE 000]: MatuX
'
Public bRain                As Boolean 'está raineando?

Public bTecho               As Boolean 'hay techo?

Public charlist(1 To 10000) As Char

'[END]'
'
'       [END]
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

Sub ConvertCPtoTP(ByVal viewPortX As Integer, _
                  ByVal viewPortY As Integer, _
                  ByRef tX As Byte, _
                  ByRef tY As Byte)
    '******************************************
    'Converts where the mouse is in the main window to a tile position. MUST be called eveytime the mouse moves.
    '******************************************
    tX = UserPos.x + viewPortX \ 32 - frmMain.renderer.ScaleWidth \ 64
    tY = UserPos.y + viewPortY \ 32 - frmMain.renderer.ScaleHeight \ 64
    Debug.Print tX; tY
End Sub

Sub DoPasosFx(ByVal charindex As Integer)

    If Not Sound Then Exit Sub

    If Not UserNavegando Then
        If Not charlist(charindex).Muerto And EstaPCarea(charindex) Then
            charlist(charindex).pie = Not charlist(charindex).pie

            If charlist(charindex).pie Then
                Call Audio.PlayWave(SND_PASOS1)
            Else
                Call Audio.PlayWave(SND_PASOS2)
            End If
        End If

    Else
        Call Audio.PlayWave(SND_NAVEGANDO)
    End If

End Sub

Sub EraseChar(ByVal charindex As Integer)

    On Error Resume Next

    '*****************************************************************
    'Erases a character from CharList and map
    '*****************************************************************

    charlist(charindex).active = 0

    'Update lastchar
    If charindex = LastChar Then

        Do Until charlist(LastChar).active = 1
            LastChar = LastChar - 1

            If LastChar = 0 Then Exit Do
        Loop

    End If

    MapData(charlist(charindex).Pos.x, charlist(charindex).Pos.y).charindex = 0
    'Remove char's dialog
    Call Dialogos.RemoveDialog(charindex)
    Call ResetCharInfo(charindex)

    'Update NumChars
    NumChars = NumChars - 1

End Sub

Function EstaPCarea(ByVal Index2 As Integer) As Boolean

    Dim x As Integer, y As Integer

    For y = UserPos.y - MinYBorder + 1 To UserPos.y + MinYBorder - 1
        For x = UserPos.x - MinXBorder + 1 To UserPos.x + MinXBorder - 1
            
            If MapData(x, y).charindex = Index2 Then
                EstaPCarea = True

                Exit Function

            End If
        
        Next x
    Next y

    EstaPCarea = False

End Function

Public Sub InitGrh(ByRef Grh As Grh, _
                   ByVal grhindex As Integer, _
                   Optional ByVal Started As Byte = 2)
    '*****************************************************************
    'Sets up a grh. MUST be done before rendering
    '*****************************************************************
    Grh.grhindex = grhindex
   
    If Started = 2 Then
        If GrhData(Grh.grhindex).NumFrames > 1 Then
            Grh.Started = 1
        Else
            Grh.Started = 0
        End If

    Else

        'Make sure the graphic can be started
        If GrhData(Grh.grhindex).NumFrames = 1 Then Started = 0
        Grh.Started = Started
    End If
   
    If Grh.Started Then
        Grh.Loops = INFINITE_LOOPS
    Else
        Grh.Loops = 0
    End If
   
    Grh.FrameCounter = 1
    Grh.Speed = GrhData(Grh.grhindex).Speed
End Sub

Function InMapBounds(ByVal x As Integer, ByVal y As Integer) As Boolean
    '*****************************************************************
    'Checks to see if a tile position is in the maps bounds
    '*****************************************************************

    If x < XMinMapSize Or x > XMaxMapSize Or y < YMinMapSize Or y > YMaxMapSize Then
        InMapBounds = False

        Exit Function

    End If

    InMapBounds = True

End Function

Function LegalPos(ByVal x As Integer, ByVal y As Integer) As Boolean
    '*****************************************************************
    'Checks to see if a tile position is legal
    '*****************************************************************

    'Limites del mapa
    If x < MinXBorder Or x > MaxXBorder Or y < MinYBorder Or y > MaxYBorder Then
        LegalPos = False

        Exit Function

    End If

    'Tile Bloqueado?
    If MapData(x, y).Blocked = 1 Then
        LegalPos = False

        Exit Function

    End If
    
    '¿Hay un personaje?
    If MapData(x, y).charindex > 0 Then
        LegalPos = False

        Exit Function

    End If
   
    If Not UserNavegando Then
        If HayAgua(x, y) Then
            LegalPos = False

            Exit Function

        End If

    Else

        If Not HayAgua(x, y) Then
            LegalPos = False

            Exit Function

        End If
    End If
    
    LegalPos = True

End Function
Sub MakeChar(ByVal charindex As Integer, ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As Byte, ByVal x As Integer, ByVal y As Integer, ByVal Arma As Integer, ByVal Escudo As Integer, ByVal Casco As Integer)
On Error Resume Next
    'Apuntamos al ultimo Char
    If charindex > LastChar Then LastChar = charindex
    
    With charlist(charindex)
        'If the char wasn't allready active (we are rewritting it) don't increase char count
        If .active = 0 Then _
            NumChars = NumChars + 1
        
        If Arma = 0 Then Arma = 2
        If Escudo = 0 Then Escudo = 2
        If Casco = 0 Then Casco = 2
        
        .iHead = Head
        .iBody = Body
        .Head = HeadData(Head)
        .Body = BodyData(Body)
        .Arma = WeaponAnimData(Arma)
        
        .Escudo = ShieldAnimData(Escudo)
        .Casco = CascoAnimData(Casco)
        
        .Heading = Heading
        
        'Reset moving stats
        .Moving = 0
        .MoveOffsetX = 0
        .MoveOffsetY = 0
        
        'Update position
        .Pos.x = x
        .Pos.y = y
        
        'Make active
        .active = 1
    End With
    
    'Plot on map
    MapData(x, y).charindex = charindex
End Sub

Sub MoveScreen(ByVal nHeading As E_Heading)

    '******************************************
    'Starts the screen moving in a direction
    '******************************************
    Dim x  As Integer

    Dim y  As Integer

    Dim tX As Integer

    Dim tY As Integer

    'Figure out which way to move
    Select Case nHeading

        Case E_Heading.NORTH
            y = -1

        Case E_Heading.EAST
            x = 1

        Case E_Heading.SOUTH
            y = 1
    
        Case E_Heading.WEST
            x = -1
        
    End Select

    'Fill temp pos
    tX = UserPos.x + x
    tY = UserPos.y + y

    'Check to see if its out of bounds
    If tX < MinXBorder Or tX > MaxXBorder Or tY < MinYBorder Or tY > MaxYBorder Then

        Exit Sub

    Else
        'Start moving... MainLoop does the rest
        AddtoUserPos.x = x
        UserPos.x = tX
        AddtoUserPos.y = y
        UserPos.y = tY
        UserMoving = 1
   
    End If

End Sub

Sub ResetCharInfo(ByVal charindex As Integer)

    charlist(charindex).active = 0
    charlist(charindex).Criminal = 0
    charlist(charindex).FxIndex = 0
    charlist(charindex).invisible = False
    charlist(charindex).Moving = 0
    charlist(charindex).Muerto = False
    charlist(charindex).Nombre = ""
    charlist(charindex).pie = False
    charlist(charindex).Pos.x = 0
    charlist(charindex).Pos.y = 0
    charlist(charindex).UsandoArma = False

End Sub



