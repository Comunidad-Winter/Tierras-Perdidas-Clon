VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDX8Engine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Rem Mannakia .. Parituclas ORE 1.0.
 
Private Type Particle
    friction As Single
    x As Single
    y As Single
    vector_x As Single
    vector_y As Single
    angle As Single
    Grh As Grh
    alive_counter As Long
    X1 As Integer
    X2 As Integer
    Y1 As Integer
    Y2 As Integer
    vecx1 As Integer
    vecx2 As Integer
    vecy1 As Integer
    vecy2 As Integer
    life1 As Long
    life2 As Long
    fric As Integer
    spin_speedL As Single
    spin_speedH As Single
    gravity As Boolean
    grav_strength As Long
    bounce_strength As Long
    spin As Boolean
    XMove As Boolean
    YMove As Boolean
    move_x1 As Integer
    move_x2 As Integer
    move_y1 As Integer
    move_y2 As Integer
    rgb_list(0 To 3) As Long
    grh_resize As Boolean
    grh_resizex As Integer
    grh_resizey As Integer
End Type
 
Dim base_tile_size As Integer
 
Private Type Stream
    Name As String
    NumOfParticles As Long
    NumGrhs As Long
    id As Long
    X1 As Long
    Y1 As Long
    X2 As Long
    Y2 As Long
    angle As Long
    vecx1 As Long
    vecx2 As Long
    vecy1 As Long
    vecy2 As Long
    life1 As Long
    life2 As Long
    friction As Long
    spin As Byte
    spin_speedL As Single
    spin_speedH As Single
    AlphaBlend As Byte
    gravity As Byte
    grav_strength As Long
    bounce_strength As Long
    XMove As Byte
    YMove As Byte
    move_x1 As Long
    move_x2 As Long
    move_y1 As Long
    move_y2 As Long
    grh_list() As Long
    colortint(0 To 3) As RGB
   
    Speed As Single
    life_counter As Long
End Type
 
'Modified by: Ryan Cain (Onezero)
'Last modify date: 5/14/2003
Private Type particle_group
    active As Boolean
    id As Long
    map_x As Integer
    map_y As Integer
    char_index As Long
 
    frame_counter As Single
    frame_speed As Single
   
    stream_type As Byte
 
    particle_stream() As Particle
    particle_count As Long
   
    grh_index_list() As Long
    grh_index_count As Long
   
    alpha_blend As Boolean
   
    alive_counter As Long
    never_die As Boolean
   
    X1 As Integer
    X2 As Integer
    Y1 As Integer
    Y2 As Integer
    angle As Integer
    vecx1 As Integer
    vecx2 As Integer
    vecy1 As Integer
    vecy2 As Integer
    life1 As Long
    life2 As Long
    fric As Long
    spin_speedL As Single
    spin_speedH As Single
    gravity As Boolean
    grav_strength As Long
    bounce_strength As Long
    spin As Boolean
    XMove As Boolean
    YMove As Boolean
    move_x1 As Integer
    move_x2 As Integer
    move_y1 As Integer
    move_y2 As Integer
    rgb_list(0 To 3) As Long
   
    'Added by Juan Martín Sotuyo Dodero
    Speed As Single
    life_counter As Long
   
    'Added by David Justus
    grh_resize As Boolean
    grh_resizex As Integer
    grh_resizey As Integer
End Type
'Particle system

 Dim map_current As Map
Dim char_last As Long
Dim char_list() As Char
Dim char_count As Long

'Dim StreamData() As particle_group
Dim TotalStreams As Long
Dim particle_group_list() As particle_group
Public particle_group_count As Long
Dim particle_group_last As Long
Rem mannakia

Private Type decoration
    Grh As Grh
    Render_On_Top As Boolean
    subtile_pos As Byte
End Type

Private Type Map_Tile
    Grh(1 To 3) As Grh
    decoration(1 To 5) As decoration
    decoration_count As Byte
    Blocked As Boolean
    particle_group_index As Long
    char_index As Long
    light_base_value(0 To 3) As Long
    light_value(0 To 3) As Long
   
    exit_index As Long
    npc_index As Long
    item_index As Long
   
    Trigger As Byte
End Type
Private Type Map
    map_grid() As Map_Tile
    map_x_max As Long
    map_x_min As Long
    map_y_max As Long
    map_y_min As Long
    map_description As String
    'Added by Juan Martín Sotuyo Dodero
    base_light_color As Long
End Type
Private Const FVF = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR

Private Declare Function QueryPerformanceFrequency _
                Lib "kernel32" (lpFrequency As Currency) As Long

Private Declare Function QueryPerformanceCounter _
                Lib "kernel32" (lpPerformanceCount As Currency) As Long

Public FPS                    As Integer

Private FramesPerSecCounter   As Integer

Private timerElapsedTime      As Single

Private timerTicksPerFrame    As Double

Private particletimer         As Single

Public engineBaseSpeed        As Single

Private lFrameTimer           As Long

Private ScrollPixelsPerFrameX As Byte

Private ScrollPixelsPerFrameY As Byte

Private HalfWindowTileWidth   As Integer

Private HalfWindowTileHeight  As Integer

Private Enum PARTICLE_STATUS

    Alive = 0
    Dead = 1

End Enum


'Dim fxs_list() As New clsFXS
'Dim fxs_count As Integer
'Dim fxs_last As Integer

Dim tex                   As Direct3DTexture8

Dim D3DbackBuffer         As Direct3DSurface8

Dim zTarget               As Direct3DSurface8

Dim stencil               As Direct3DSurface8

Dim superTex              As Direct3DSurface8

'Sets a Grh animation to loop indefinitely.

Private Declare Function SelectObject _
                Lib "gdi32" (ByVal hDC As Long, _
                             ByVal hObject As Long) As Long

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long

Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
 
'Added by Juan Martín Sotuyo Dodero
Private Declare Function SetPixel _
                Lib "gdi32" (ByVal hDC As Long, _
                             ByVal x As Long, _
                             ByVal y As Long, _
                             ByVal crColor As Long) As Long

Private Declare Function GetPixel _
                Lib "gdi32" (ByVal hDC As Long, _
                             ByVal x As Long, _
                             ByVal y As Long) As Long
'Added by Barrin

#Const HARDCODED = False 'True ' == MÁS FPS ^^

Public Sub setup_ambient()

    'Noche 87, 61, 43
    luz_dia(0).r = 87
    luz_dia(0).g = 61
    luz_dia(0).B = 43
    luz_dia(1).r = 87
    luz_dia(1).g = 61
    luz_dia(1).B = 43
    luz_dia(2).r = 87
    luz_dia(2).g = 61
    luz_dia(2).B = 43
    luz_dia(3).r = 87
    luz_dia(3).g = 61
    luz_dia(3).B = 43
    '4 am 124,117,91
    luz_dia(4).r = 124
    luz_dia(4).g = 127
    luz_dia(4).B = 91
    '5,6 am 143,137,135
    luz_dia(5).r = 143
    luz_dia(5).g = 137
    luz_dia(5).B = 135
    luz_dia(6).r = 143
    luz_dia(6).g = 137
    luz_dia(6).B = 135
    '7 am 212,205,207
    luz_dia(7).r = 212
    luz_dia(7).g = 205
    luz_dia(7).B = 207
    luz_dia(8).r = 212
    luz_dia(8).g = 205
    luz_dia(8).B = 207
    luz_dia(9).r = 212
    luz_dia(9).g = 205
    luz_dia(9).B = 207
    luz_dia(10).r = 212
    luz_dia(10).g = 205
    luz_dia(10).B = 207
    luz_dia(11).r = 212
    luz_dia(11).g = 205
    luz_dia(11).B = 207
    luz_dia(12).r = 212
    luz_dia(12).g = 205
    luz_dia(12).B = 207
    'Dia 255, 255, 255
    luz_dia(12).r = 255
    luz_dia(12).g = 255
    luz_dia(12).B = 255
    luz_dia(13).r = 255
    luz_dia(13).g = 255
    luz_dia(13).B = 255
    'Medio Dia 255, 200, 255
    luz_dia(14).r = 255
    luz_dia(14).g = 250
    luz_dia(14).B = 255
    luz_dia(15).r = 255
    luz_dia(15).g = 240
    luz_dia(15).B = 255
    luz_dia(16).r = 255
    luz_dia(16).g = 230
    luz_dia(16).B = 255
    '17/18 0, 100, 255
    luz_dia(17).r = 230
    luz_dia(17).g = 230
    luz_dia(17).B = 255
    '18/19 0, 100, 255
    luz_dia(18).r = 230
    luz_dia(18).g = 230
    luz_dia(18).B = 255
    '19/20 156, 142, 83
    luz_dia(19).r = 156
    luz_dia(19).g = 142
    luz_dia(19).B = 83
    luz_dia(20).r = 87
    luz_dia(20).g = 61
    luz_dia(20).B = 43
    luz_dia(21).r = 87
    luz_dia(21).g = 61
    luz_dia(21).B = 43
    luz_dia(22).r = 87
    luz_dia(22).g = 61
    luz_dia(22).B = 43
    luz_dia(23).r = 87
    luz_dia(23).g = 61
    luz_dia(23).B = 43
    luz_dia(24).r = 87
    luz_dia(24).g = 61
    luz_dia(24).B = 43
            
End Sub

Private Function GetElapsedTime() As Single

    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 10/07/2002
    'Gets the time that past since the last call
    '**************************************************************
    Dim start_time    As Currency

    Static end_time   As Currency

    Static timer_freq As Currency

    'Get the timer frequency
    If timer_freq = 0 Then
        QueryPerformanceFrequency timer_freq
    End If
    
    'Get current time
    Call QueryPerformanceCounter(start_time)
    
    'Calculate elapsed time
    GetElapsedTime = (start_time - end_time) / timer_freq * 1000
    
    'Get next end time
    Call QueryPerformanceCounter(end_time)
End Function

Function MakeVector(ByVal x As Single, ByVal y As Single, ByVal Z As Single) As D3DVECTOR
    '*****************************************************
    '****** Coded by Menduz (lord.yo.wo@gmail.com) *******
    '*****************************************************
    MakeVector.x = x
    MakeVector.y = y
    MakeVector.Z = Z
End Function

Public Sub Engine_Init()
    '*****************************************************
    '****** Coded by Menduz (lord.yo.wo@gmail.com) *******
    '*****************************************************
    'On Error GoTo ErrHandler:

    Dim DispMode   As D3DDISPLAYMODE

    Dim DispModeBK As D3DDISPLAYMODE

    Dim D3DWindow  As D3DPRESENT_PARAMETERS
    
    Set SurfaceDB = New clsSurfaceManDynDX8
    
    Set dX = New DirectX8
    Set D3D = dX.Direct3DCreate()
    Set D3DX = New D3DX8
    
    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode
    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispModeBK
    
    With D3DWindow
        .Windowed = True
        .SwapEffect = D3DSWAPEFFECT_COPY
        .BackBufferFormat = DispMode.Format
        .BackBufferWidth = frmMain.renderer.ScaleWidth
        .BackBufferHeight = frmMain.renderer.ScaleHeight
        .EnableAutoDepthStencil = 1
        .AutoDepthStencilFormat = D3DFMT_D16
        .hDeviceWindow = frmMain.renderer.hWnd
    End With

    DispMode.Format = D3DFMT_X8R8G8B8

    If D3D.CheckDeviceFormat(0, D3DDEVTYPE_HAL, DispMode.Format, 0, D3DRTYPE_TEXTURE, D3DFMT_A8R8G8B8) = D3D_OK Then

        Dim Caps8 As D3DCAPS8

        D3D.GetDeviceCaps 0, D3DDEVTYPE_HAL, Caps8

        If (Caps8.TextureOpCaps And D3DTEXOPCAPS_DOTPRODUCT3) = D3DTEXOPCAPS_DOTPRODUCT3 Then
        Else
            DispMode.Format = DispModeBK.Format
        End If

    Else
        DispMode.Format = DispModeBK.Format
    End If

    Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.renderer.hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, D3DWindow)
                                                            
    HalfWindowTileHeight = (frmMain.renderer.ScaleHeight / 32) \ 2
    HalfWindowTileWidth = (frmMain.renderer.ScaleWidth / 32) \ 2
    
    TileBufferSize = 9
    
    D3DDevice.SetVertexShader FVF
    
    '//Transformed and lit vertices dont need lighting
    '   so we disable it...
    D3DDevice.SetRenderState D3DRS_LIGHTING, False
    
    D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, True
    
    Call SurfaceDB.Init(D3DX, D3DDevice, General_Get_Free_Ram_Bytes)

    engineBaseSpeed = 0.015
    
    ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
    
    'Set FPS value to 60 for startup
    FPS = 60
    FramesPerSecCounter = 60
    
    ScrollPixelsPerFrameX = 9
    ScrollPixelsPerFrameY = 9
    
    UserPos.x = 50
    UserPos.y = 50
    
    MinXBorder = XMinMapSize + (frmMain.renderer.ScaleWidth / 64)
    MaxXBorder = XMaxMapSize - (frmMain.renderer.ScaleWidth / 64)
    MinYBorder = YMinMapSize + (frmMain.renderer.ScaleHeight / 64)
    MaxYBorder = YMaxMapSize - (frmMain.renderer.ScaleHeight / 64)
    
    'partículas
    D3DDevice.SetRenderState D3DRS_POINTSIZE, Engine_FToDW(2)
    D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
    D3DDevice.SetRenderState D3DRS_POINTSPRITE_ENABLE, 1
    D3DDevice.SetRenderState D3DRS_POINTSCALE_ENABLE, 0
    'motion blur
    Set D3DbackBuffer = D3DDevice.GetRenderTarget
    Set zTarget = D3DDevice.GetDepthStencilSurface
    Set stencil = D3DDevice.CreateDepthStencilSurface(800, 600, D3DFMT_D16, D3DMULTISAMPLE_NONE)
    Set tex = D3DX.CreateTexture(D3DDevice, 0, 0, 1, D3DUSAGE_RENDERTARGET, D3DFMT_X8R8G8B8, D3DPOOL_DEFAULT)
    Set superTex = tex.GetSurfaceLevel(0)
    'bump mapping

    Exit Sub

ErrHandler:
    Debug.Print "Error Number Returned: " & Err.Number
End Sub

Public Sub Engine_Deinit()
    Erase MapData
    Erase charlist
    Set D3DDevice = Nothing
    Set D3D = Nothing
    Set dX = Nothing

    End

End Sub

Public Sub Engine_ActFPS()

    If GetTickCount - lFrameTimer > 1000 Then
        FPS = FramesPerSecCounter
        FramesPerSecCounter = 0
        lFrameTimer = GetTickCount
    End If

End Sub

Public Sub Draw_GrhIndex(ByVal Grh_Index As Integer, _
                         ByVal x As Integer, _
                         ByVal y As Integer)

    If Grh_Index <= 0 Then Exit Sub

    Dim rgb_list(3) As Long
    
    rgb_list(0) = D3DColorXRGB(255, 255, 255)
    rgb_list(1) = D3DColorXRGB(255, 255, 255)
    rgb_list(2) = D3DColorXRGB(255, 255, 255)
    rgb_list(3) = D3DColorXRGB(255, 255, 255)
    
    Device_Box_Textured_Render Grh_Index, x, y, GrhData(Grh_Index).pixelWidth, GrhData(Grh_Index).pixelHeight, rgb_list, GrhData(Grh_Index).sX, GrhData(Grh_Index).sY
End Sub

Private Sub Draw_Grh(ByRef Grh As Grh, _
                     ByVal x As Integer, _
                     ByVal y As Integer, _
                     ByVal center As Byte, _
                     ByVal Animate As Byte, _
                     Optional ByVal alpha As Boolean, _
                     Optional ByVal map_x As Byte = 1, _
                     Optional ByVal map_y As Byte = 1, _
                     Optional ByVal angle As Single)

    Dim CurrentGrhIndex As Integer

    If Grh.grhindex = 0 Then Exit Sub
    If Animate Then
        If Grh.Started = 1 Then
            Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * GrhData(Grh.grhindex).NumFrames / Grh.Speed)

            If Grh.FrameCounter > GrhData(Grh.grhindex).NumFrames Then
                Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.grhindex).NumFrames) + 1
                
                If Grh.Loops <> -1 Then
                    If Grh.Loops > 0 Then
                        Grh.Loops = Grh.Loops - 1
                    Else
                        Grh.Started = 0
                    End If
                End If
            End If
        End If
    End If
    
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(Grh.grhindex).Frames(Grh.FrameCounter)

    'Center Grh over X,Y pos
    If center Then
        If GrhData(CurrentGrhIndex).TileWidth <> 1 Then
            x = x - Int(GrhData(CurrentGrhIndex).TileWidth * (32 \ 2)) + 32 \ 2
        End If

        If GrhData(Grh.grhindex).TileHeight <> 1 Then
            y = y - Int(GrhData(CurrentGrhIndex).TileHeight * 32) + 32
        End If
    End If

    Device_Box_Textured_Render CurrentGrhIndex, x, y, GrhData(CurrentGrhIndex).pixelWidth, GrhData(CurrentGrhIndex).pixelHeight, MapData(map_x, map_y).light_value, GrhData(CurrentGrhIndex).sX, GrhData(CurrentGrhIndex).sY, alpha, angle
    'exits:

End Sub

Public Sub Render()
    '*****************************************************
    '****** Coded by Menduz (lord.yo.wo@gmail.com) *******
    '*****************************************************

    D3DDevice.BeginScene

    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0

    ShowNextFrame

    frmMain.FPSMain.Caption = "FPS: " & FPS
    engine.Engine_ActFPS
    Call Dialogos.Render
    Call DialogosClanes.Draw
    D3DDevice.EndScene
    D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0

    FramesPerSecCounter = FramesPerSecCounter + 1
    timerElapsedTime = GetElapsedTime()
    timerTicksPerFrame = timerElapsedTime * engineBaseSpeed
    particletimer = timerElapsedTime * 0.05
End Sub

Sub ShowNextFrame()

    Static OffsetCounterX As Single

    Static OffsetCounterY As Single

    If UserMoving Then

        '****** Move screen Left and Right if needed ******
        If AddtoUserPos.x <> 0 Then
            OffsetCounterX = OffsetCounterX - ScrollPixelsPerFrameX * AddtoUserPos.x * timerTicksPerFrame

            If Abs(OffsetCounterX) >= Abs(32 * AddtoUserPos.x) Then
                OffsetCounterX = 0
                AddtoUserPos.x = 0
                UserMoving = False
            End If
        End If
            
        '****** Move screen Up and Down if needed ******
        If AddtoUserPos.y <> 0 Then
            OffsetCounterY = OffsetCounterY - ScrollPixelsPerFrameY * AddtoUserPos.y * timerTicksPerFrame

            If Abs(OffsetCounterY) >= Abs(32 * AddtoUserPos.y) Then
                OffsetCounterY = 0
                AddtoUserPos.y = 0
                UserMoving = False
            End If
        End If
    End If

    Call RenderScreen(UserPos.x - AddtoUserPos.x, UserPos.y - AddtoUserPos.y, OffsetCounterX, OffsetCounterY)
End Sub

Sub RenderScreen(ByVal tilex As Integer, _
                 ByVal tiley As Integer, _
                 ByVal PixelOffsetX As Integer, _
                 ByVal PixelOffsetY As Integer)

    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 8/14/2007
    'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
    'Renders everything to the viewport
    '**************************************************************
    Dim y                As Integer     'Keeps track of where on map we are

    Dim x                As Integer     'Keeps track of where on map we are

    Dim screenminY       As Integer  'Start Y pos on current screen

    Dim screenmaxY       As Integer  'End Y pos on current screen

    Dim screenminX       As Integer  'Start X pos on current screen

    Dim screenmaxX       As Integer  'End X pos on current screen

    Dim minY             As Integer  'Start Y pos on current map

    Dim maxY             As Integer  'End Y pos on current map

    Dim minX             As Integer  'Start X pos on current map

    Dim maxX             As Integer  'End X pos on current map

    Dim ScreenX          As Integer  'Keeps track of where to place tile on screen

    Dim ScreenY          As Integer  'Keeps track of where to place tile on screen

    Dim minXOffset       As Integer

    Dim minYOffset       As Integer

    Dim PixelOffsetXTemp As Integer 'For centering grhs

    Dim PixelOffsetYTemp As Integer 'For centering grhs

    Dim CurrentGrhIndex  As Integer

    'Figure out Ends and Starts of screen
    screenminY = tiley - HalfWindowTileHeight
    screenmaxY = tiley + HalfWindowTileHeight
    screenminX = tilex - HalfWindowTileWidth
    screenmaxX = tilex + HalfWindowTileWidth
    
    minY = screenminY - TileBufferSize
    maxY = screenmaxY + TileBufferSize
    minX = screenminX - TileBufferSize
    maxX = screenmaxX + TileBufferSize
    
    'Make sure mins and maxs are allways in map bounds
    If minY < XMinMapSize Then
        minYOffset = YMinMapSize - minY
        minY = YMinMapSize
    End If
    
    If maxY > YMaxMapSize Then maxY = YMaxMapSize
    
    If minX < XMinMapSize Then
        minXOffset = XMinMapSize - minX
        minX = XMinMapSize
    End If
    
    If maxX > XMaxMapSize Then maxX = XMaxMapSize
    
    'If we can, we render around the view area to make it smoother
    If screenminY > YMinMapSize Then
        screenminY = screenminY - 1
    Else
        screenminY = 1
        ScreenY = 1
    End If
    
    If screenmaxY < YMaxMapSize Then screenmaxY = screenmaxY + 1
    
    If screenminX > XMinMapSize Then
        screenminX = screenminX - 1
    Else
        screenminX = 1
        ScreenX = 1
    End If
    
    If screenmaxX < XMaxMapSize Then screenmaxX = screenmaxX + 1
    
    'Draw floor layer
    For y = screenminY To screenmaxY
        For x = screenminX To screenmaxX
            'Layer 1 **********************************
            
            #If HARDCODED = True Then

                If MapData(x, y).Graphic(1).Started = 1 Then
                    MapData(x, y).Graphic(1).FrameCounter = MapData(x, y).Graphic(1).FrameCounter + (timerElapsedTime * GrhData(MapData(x, y).Graphic(1).grhindex).NumFrames / MapData(x, y).Graphic(1).Speed)

                    If MapData(x, y).Graphic(1).FrameCounter > GrhData(MapData(x, y).Graphic(1).grhindex).NumFrames Then
                        MapData(x, y).Graphic(1).FrameCounter = (MapData(x, y).Graphic(1).FrameCounter Mod GrhData(MapData(x, y).Graphic(1).grhindex).NumFrames) + 1
                            
                        If MapData(x, y).Graphic(1).Loops <> -1 Then
                            If MapData(x, y).Graphic(1).Loops > 0 Then
                                MapData(x, y).Graphic(1).Loops = MapData(x, y).Graphic(1).Loops - 1
                            Else
                                MapData(x, y).Graphic(1).Started = 0
                            End If
                        End If
                    End If
                End If

                CurrentGrhIndex = GrhData(MapData(x, y).Graphic(1).grhindex).Frames(MapData(x, y).Graphic(1).FrameCounter)

                Device_Box_Textured_Render CurrentGrhIndex, (ScreenX - 1) * 32 + PixelOffsetX, (ScreenY - 1) * 32 + PixelOffsetY, GrhData(CurrentGrhIndex).pixelWidth, GrhData(CurrentGrhIndex).pixelHeight, MapData(x, y).light_value, GrhData(CurrentGrhIndex).sX, GrhData(CurrentGrhIndex).sY, False, 0
            #Else
                Call Draw_Grh(MapData(x, y).Graphic(1), (ScreenX - 1) * 32 + PixelOffsetX, (ScreenY - 1) * 32 + PixelOffsetY, 0, 1, , x, y)

                If MapData(x, y).Graphic(2).grhindex <> 0 Then
                    Call Draw_Grh(MapData(x, y).Graphic(2), (ScreenX - 1) * 32 + PixelOffsetX, (ScreenY - 1) * 32 + PixelOffsetY, 1, 1, , x, y)
                End If

            #End If
            '******************************************
            ScreenX = ScreenX + 1
        Next x

        'Reset ScreenX to original value and increment ScreenY
        ScreenX = ScreenX - x + screenminX
        ScreenY = ScreenY + 1
    Next y
    
    'Draw floor layer 2
    '    ScreenY = minYOffset
    '    For Y = screenminY To screenmaxY
    '        ScreenX = minXOffset
    '        For X = screenminX To screenmaxX
    'Layer 2 **********************************
    '                If MapData(X, Y).Graphic(2).grhindex <> 0 Then
    '#If HARDCODED = True Then
    '                    If MapData(X, Y).Graphic(2).Started = 1 Then
    '                        MapData(X, Y).Graphic(2).FrameCounter = MapData(X, Y).Graphic(2).FrameCounter + (timerElapsedTime * GrhData(MapData(X, Y).Graphic(2).grhindex).NumFrames / MapData(X, Y).Graphic(2).Speed)
    '                        If MapData(X, Y).Graphic(2).FrameCounter > GrhData(MapData(X, Y).Graphic(2).grhindex).NumFrames Then
    '                            MapData(X, Y).Graphic(2).FrameCounter = (MapData(X, Y).Graphic(2).FrameCounter Mod GrhData(MapData(X, Y).Graphic(2).grhindex).NumFrames) + 1
    '
    '                            If MapData(X, Y).Graphic(2).Loops <> -1 Then
    '                                If MapData(X, Y).Graphic(2).Loops > 0 Then
    '                                    MapData(X, Y).Graphic(2).Loops = MapData(X, Y).Graphic(2).Loops - 1
    '                                Else
    '                                    MapData(X, Y).Graphic(2).Started = 0
    '                                End If
    '                            End If
    '                        End If
    '                    End If
    '
    '                CurrentGrhIndex = GrhData(MapData(X, Y).Graphic(2).grhindex).Frames(MapData(X, Y).Graphic(2).FrameCounter)
    '
    '                offx = 0
    '                offy = 0
    '                If GrhData(CurrentGrhIndex).TileWidth <> 1 Then _
    '                    offx = -Int(GrhData(CurrentGrhIndex).TileWidth * (32 \ 2)) + 32 \ 2
    '                If GrhData(MapData(X, Y).Graphic(2).grhindex).TileHeight <> 1 Then _
    '                    offy = -Int(GrhData(CurrentGrhIndex).TileHeight * 32) + 32
    '
    '                Device_Box_Textured_Render CurrentGrhIndex, _
    '                    (ScreenX - 1) * 32 + PixelOffsetX + offx, (ScreenY - 1) * 32 + PixelOffsetY + offy, _
    '                    GrhData(CurrentGrhIndex).pixelWidth, GrhData(CurrentGrhIndex).pixelHeight, _
    '                    MapData(X, Y).light_value, _
    '                    GrhData(CurrentGrhIndex).sX, GrhData(CurrentGrhIndex).sY, _
    '                    False _
    '                    , 0
    '#Else
    ''                    Call Draw_Grh(MapData(X, Y).Graphic(2), _
    '                            (ScreenX - 1) * 32 + PixelOffsetX, _
    '                            (ScreenY - 1) * 32 + PixelOffsetY, _
    '                            1, 1, , X, Y)
    '#End If
    ''                End If
    '
    ''            ScreenX = ScreenX + 1
    ''        Next X'
    '
    '        'Reset ScreenX to original value and increment ScreenY
    '        'ScreenX = ScreenX - X + screenminX
    '        'ScreenY = ScreenY + 1
    '    'Next Y
    
    ScreenY = minYOffset - TileBufferSize

    For y = minY To maxY
        ScreenX = minXOffset - TileBufferSize

        For x = minX To maxX
            PixelOffsetXTemp = ScreenX * 32 + PixelOffsetX
            PixelOffsetYTemp = ScreenY * 32 + PixelOffsetY

            With MapData(x, y)
                '******************************************

                'Object Layer **********************************
                If .ObjGrh.grhindex <> 0 Then
                    '#If HARDCODED = True Then
                    '                    CurrentGrhIndex = GrhData(.ObjGrh.grhindex).Frames(.ObjGrh.FrameCounter)
                    '                    offx = 0
                    '                    offy = 0
                    '                    If GrhData(CurrentGrhIndex).TileWidth <> 1 Then _
                    '                        offx = -Int(GrhData(CurrentGrhIndex).TileWidth * (32 \ 2)) + 32 \ 2
                    '                    If GrhData(.ObjGrh.grhindex).TileHeight <> 1 Then _
                    '                        offy = -Int(GrhData(CurrentGrhIndex).TileHeight * 32) + 32
                    '
                    '                    Device_Box_Textured_Render CurrentGrhIndex, _
                    '                        PixelOffsetXTemp + offx, PixelOffsetYTemp + offx, _
                    '                        GrhData(CurrentGrhIndex).pixelWidth, GrhData(CurrentGrhIndex).pixelHeight, _
                    '                        MapData(X, Y).light_value, _
                    '                        GrhData(CurrentGrhIndex).sX, GrhData(CurrentGrhIndex).sY, _
                    '                        False _
                    '                        , 0
                    '#Else
                    Call Draw_Grh(.ObjGrh, PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, , x, y)
                    '#End If
                End If

                '***********************************************
                'If .particle_group > 0 Then
                '    Particles_Update .particle_group, PixelOffsetYTemp, PixelOffsetXTemp
                '
                '             D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
                '             D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
                '
                '             Dim tmplng As Long
                '             D3DDevice.SetTexture 0, SurfaceDB.GetTexture(GrhData(particle_group_list(.particle_group).texture).FileNum, tmplng) 'Textures(particle_group_list(.particle_group).Texture)
                '             D3DDevice.SetRenderState D3DRS_POINTSIZE, Engine_FToDW(particle_group_list(.particle_group).size)
                '             D3DDevice.DrawPrimitiveUP D3DPT_POINTLIST, particle_group_list(.particle_group).nParticles, particle_group_list(.particle_group).PrtVertList(0), Len(particle_group_list(.particle_group).PrtVertList(0))
                '             D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
                '             D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
                '
                '         End If
                
                'Char layer ************************************
                If .charindex <> 0 Then
                    Call Char_Render(.charindex, PixelOffsetXTemp, PixelOffsetYTemp, x, y)
                End If

                '*************************************************
                
                'Layer 3 *****************************************
                If .Graphic(3).grhindex <> 0 Then
                    '#If HARDCODED = True Then
                    '                    If .Graphic(3).Started = 1 Then
                    '                        .Graphic(3).FrameCounter = .Graphic(3).FrameCounter + (timerElapsedTime * GrhData(.Graphic(3).grhindex).NumFrames / .Graphic(3).Speed)
                    '                        If .Graphic(3).FrameCounter > GrhData(.Graphic(3).grhindex).NumFrames Then
                    '                            .Graphic(3).FrameCounter = (.Graphic(3).FrameCounter Mod GrhData(.Graphic(3).grhindex).NumFrames) + 1
                    '
                    '                            If .Graphic(3).Loops <> -1 Then
                    '                                If .Graphic(3).Loops > 0 Then
                    '                                    .Graphic(3).Loops = .Graphic(3).Loops - 1
                    '                                Else
                    '                                    .Graphic(3).Started = 0
                    '                                End If
                    '                            End If
                    '                        End If
                    '                    End If
                    '
                    '                CurrentGrhIndex = GrhData(.Graphic(3).grhindex).Frames(.Graphic(3).FrameCounter)
                    '
                    '                offx = 0
                    '                offy = 0
                    '                If GrhData(CurrentGrhIndex).TileWidth <> 1 Then _
                    '                    offx = -Int(GrhData(CurrentGrhIndex).TileWidth * (32 \ 2)) + 32 \ 2
                    '                If GrhData(.Graphic(3).grhindex).TileHeight <> 1 Then _
                    '                    offy = -Int(GrhData(CurrentGrhIndex).TileHeight * 32) + 32
                    '
                    '                Device_Box_Textured_Render CurrentGrhIndex, _
                    '                    PixelOffsetXTemp + offx, PixelOffsetYTemp + offy, _
                    '                    GrhData(CurrentGrhIndex).pixelWidth, GrhData(CurrentGrhIndex).pixelHeight, _
                    '                    MapData(X, Y).light_value, _
                    '                    GrhData(CurrentGrhIndex).sX, GrhData(CurrentGrhIndex).sY, _
                    '                    False _
                    '                    , 0
                    '#Else
                    Call Draw_Grh(.Graphic(3), PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, , x, y)
                    '#End If
                End If

                '************************************************

            End With

            ScreenX = ScreenX + 1
        Next x

        ScreenY = ScreenY + 1
    Next y

    ScreenY = minYOffset - 5
    'render ´particulas
    
                    ScreenY = minYOffset - TileBufferSize
    For y = minY To maxY
        ScreenX = minXOffset - TileBufferSize
        For x = minX To maxX
                If MapData(x, y).particle_group_index Then
                    Particle_Group_Render MapData(x, y).particle_group_index, ScreenX * 32 + PixelOffsetX, ScreenY * 32 + PixelOffsetY  '+ (16)
                End If
            ScreenX = ScreenX + 1
        Next x
        ScreenY = ScreenY + 1
    Next y
    'render particulas
                    
                    
                    
    If Not bTecho Then
        'Draw blocked tiles and grid
        ScreenY = minYOffset - TileBufferSize

        For y = minY To maxY
            ScreenX = minXOffset - TileBufferSize

            For x = minX To maxX

                'Layer 4 **********************************
                If MapData(x, y).Graphic(4).grhindex Then
                    '#If HARDCODED = True Then
                    '                CurrentGrhIndex = GrhData(MapData(X, Y).Graphic(4).grhindex).Frames(MapData(X, Y).Graphic(4).FrameCounter)
                    '                offx = 0
                    '                offy = 0
                    '                If GrhData(CurrentGrhIndex).TileWidth <> 1 Then _
                    '                    offx = -Int(GrhData(CurrentGrhIndex).TileWidth * (32 \ 2)) + 32 \ 2
                    '                If GrhData(MapData(X, Y).Graphic(4).grhindex).TileHeight <> 1 Then _
                    '                    offy = -Int(GrhData(CurrentGrhIndex).TileHeight * 32) + 32
                    '
                    '                Device_Box_Textured_Render CurrentGrhIndex, _
                    '                    (ScreenX - 1) * 32 + PixelOffsetX + offx, (ScreenY - 1) * 32 + PixelOffsetY + offy, _
                    '                    GrhData(CurrentGrhIndex).pixelWidth, GrhData(CurrentGrhIndex).pixelHeight, _
                    '                    MapData(X, Y).light_value, _
                    '                    GrhData(CurrentGrhIndex).sX, GrhData(CurrentGrhIndex).sY, _
                    '                    False _
                    '                    , 0
                    '#Else
                    Call Draw_Grh(MapData(x, y).Graphic(4), ScreenX * 32 + PixelOffsetX, ScreenY * 32 + PixelOffsetY, 1, 1, True, x, y)
                    '#End If
                End If

                '**********************************
                
                ScreenX = ScreenX + 1
            Next x

            ScreenY = ScreenY + 1
        Next y

    End If
    
    Call DibujarCartel
    
End Sub

Private Function Geometry_Create_TLVertex(ByVal x As Single, _
                                          ByVal y As Single, _
                                          ByVal Z As Single, _
                                          ByVal rhw As Single, _
                                          ByVal color As Long, _
                                          ByVal Specular As Long, _
                                          tu As Single, _
                                          ByVal tv As Single) As TLVERTEX
    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 10/07/2002
    '**************************************************************
    Geometry_Create_TLVertex.x = x
    Geometry_Create_TLVertex.y = y
    Geometry_Create_TLVertex.Z = Z
    Geometry_Create_TLVertex.rhw = rhw
    Geometry_Create_TLVertex.color = color
    Geometry_Create_TLVertex.Specular = Specular
    Geometry_Create_TLVertex.tu = tu
    Geometry_Create_TLVertex.tv = tv
End Function

Private Sub Geometry_Create_Box(ByRef verts() As TLVERTEX, _
                                ByRef dest As RECT, _
                                ByRef src As RECT, _
                                ByRef rgb_list() As Long, _
                                Optional ByRef Textures_Width As Long, _
                                Optional ByRef Textures_Height As Long, _
                                Optional ByVal angle As Single)

    '**************************************************************
    'Author: Aaron Perkins
    'Modified by Juan Martín Sotuyo Dodero
    'Last Modify Date: 11/17/2002
    '
    ' * v1      * v3
    ' |\        |
    ' |  \      |
    ' |    \    |
    ' |      \  |
    ' |        \|
    ' * v0      * v2
    '**************************************************************
    Dim x_center    As Single

    Dim y_center    As Single

    Dim radius      As Single

    Dim x_Cor       As Single

    Dim y_Cor       As Single

    Dim left_point  As Single

    Dim right_point As Single

    Dim temp        As Single
    
    If angle > 0 Then
        'Center coordinates on screen of the square
        x_center = dest.Left + (dest.Right - dest.Left) / 2
        y_center = dest.Top + (dest.bottom - dest.Top) / 2
        
        'Calculate radius
        radius = Sqr((dest.Right - x_center) ^ 2 + (dest.bottom - y_center) ^ 2)
        
        'Calculate left and right points
        temp = (dest.Right - x_center) / radius
        right_point = Atn(temp / Sqr(-temp * temp + 1))
        left_point = PI - right_point
    End If
    
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Left
        y_Cor = dest.bottom
    Else
        x_Cor = x_center + Cos(-left_point - angle) * radius
        y_Cor = y_center - Sin(-left_point - angle) * radius
    End If
    
    '0 - Bottom left vertex
    If Textures_Width And Textures_Height Then
        verts(0) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(0), 0, src.Left / Textures_Width, (src.bottom + 1) / Textures_Height)
    Else
        verts(0) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(0), 0, 0, 0)
    End If

    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Left
        y_Cor = dest.Top
    Else
        x_Cor = x_center + Cos(left_point - angle) * radius
        y_Cor = y_center - Sin(left_point - angle) * radius
    End If
    
    '1 - Top left vertex
    If Textures_Width And Textures_Height Then
        verts(1) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(1), 0, src.Left / Textures_Width, src.Top / Textures_Height)
    Else
        verts(1) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(1), 0, 0, 1)
    End If

    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Right
        y_Cor = dest.bottom
    Else
        x_Cor = x_center + Cos(-right_point - angle) * radius
        y_Cor = y_center - Sin(-right_point - angle) * radius
    End If
    
    '2 - Bottom right vertex
    If Textures_Width And Textures_Height Then
        verts(2) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(2), 0, (src.Right + 1) / Textures_Width, (src.bottom + 1) / Textures_Height)
    Else
        verts(2) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(2), 0, 1, 0)
    End If

    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Right
        y_Cor = dest.Top
    Else
        x_Cor = x_center + Cos(right_point - angle) * radius
        y_Cor = y_center - Sin(right_point - angle) * radius
    End If
    
    '3 - Top right vertex
    If Textures_Width And Textures_Height Then
        verts(3) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(3), 0, (src.Right + 1) / Textures_Width, src.Top / Textures_Height)
    Else
        verts(3) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(3), 0, 1, 1)
    End If

End Sub

Public Sub Device_Box_Textured_Render(ByVal grhindex As Long, _
                                      ByVal dest_x As Integer, _
                                      ByVal dest_y As Integer, _
                                      ByVal src_width As Integer, _
                                      ByVal src_height As Integer, _
                                      ByRef rgb_list() As Long, _
                                      ByVal src_x As Integer, _
                                      ByVal src_y As Integer, _
                                      Optional ByVal alpha_blend As Boolean, _
                                      Optional ByVal angle As Single)

    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero
    'Last Modify Date: 2/12/2004
    'Just copies the Textures
    '**************************************************************
    Static src_rect            As RECT

    Static dest_rect           As RECT

    Static temp_verts(3)       As TLVERTEX

    Static d3dTextures         As D3D8Textures

    Static light_value(0 To 3) As Long
    
    If grhindex = 0 Then Exit Sub
    Set d3dTextures.texture = SurfaceDB.GetTexture(GrhData(grhindex).FileNum, d3dTextures.texwidth, d3dTextures.texheight)
    
    light_value(0) = rgb_list(0)
    light_value(1) = rgb_list(1)
    light_value(2) = rgb_list(2)
    light_value(3) = rgb_list(3)
    
    'If Not char_current_blind Then
    If (light_value(0) = 0) Then light_value(0) = base_light
    If (light_value(1) = 0) Then light_value(1) = base_light
    If (light_value(2) = 0) Then light_value(2) = base_light
    If (light_value(3) = 0) Then light_value(3) = base_light
    'Else
    '    light_value(0) = &HFFFFFFFF 'blind_color
    '    light_value(1) = &HFFFFFFFF 'blind_color
    '    light_value(2) = &HFFFFFFFF 'blind_color
    '    light_value(3) = &HFFFFFFFF 'blind_color
    'End If
        
    'Set up the source rectangle
    With src_rect
        .bottom = src_y + src_height
        .Left = src_x
        .Right = src_x + src_width
        .Top = src_y
    End With
                
    'Set up the destination rectangle
    With dest_rect
        .bottom = dest_y + src_height
        .Left = dest_x
        .Right = dest_x + src_width
        .Top = dest_y
    End With
    
    'Set up the TempVerts(3) vertices
    Geometry_Create_Box temp_verts(), dest_rect, src_rect, light_value(), d3dTextures.texwidth, d3dTextures.texheight, angle
    
    'Set Textures
    D3DDevice.SetTexture 0, d3dTextures.texture
    
    If alpha_blend Then
        'Set Rendering for alphablending
        D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    End If
    
    'Draw the triangles that make up our square Textures
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, temp_verts(0), Len(temp_verts(0))
    
    If alpha_blend Then
        'Set Rendering for colokeying
        D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    End If

End Sub

Private Sub Char_Render(ByVal charindex As Long, _
                        ByVal PixelOffsetX As Integer, _
                        ByVal PixelOffsetY As Integer, _
                        ByVal x As Byte, _
                        ByVal y As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 12/03/04
    'Draw char's to screen without offcentering them
    '***************************************************
    Dim moved As Boolean

    Dim Pos   As Integer

    Dim line  As String

    Dim color As Long
    
    With charlist(charindex)

        If .Moving Then

            'If needed, move left and right
            If .scrollDirectionX <> 0 Then
                .MoveOffsetX = .MoveOffsetX + ScrollPixelsPerFrameX * Sgn(.scrollDirectionX) * timerTicksPerFrame
                
                'Start animations
                'TODO : Este parche es para evita los uncornos exploten al moverse!! REVER!!!
                If .Body.Walk(.Heading).Speed > 0 Then .Body.Walk(.Heading).Started = 1
                .Arma.WeaponWalk(.Heading).Started = 1
                .Escudo.ShieldWalk(.Heading).Started = 1
                
                'Char moved
                moved = True
                
                'Check if we already got there
                If (Sgn(.scrollDirectionX) = 1 And .MoveOffsetX >= 0) Or (Sgn(.scrollDirectionX) = -1 And .MoveOffsetX <= 0) Then
                    .MoveOffsetX = 0
                    .scrollDirectionX = 0
                End If
            End If
            
            'If needed, move up and down
            If .scrollDirectionY <> 0 Then
                .MoveOffsetY = .MoveOffsetY + ScrollPixelsPerFrameY * Sgn(.scrollDirectionY) * timerTicksPerFrame
                
                'Start animations
                'TODO : Este parche es para evita los uncornos exploten al moverse!! REVER!!!
                If .Body.Walk(.Heading).Speed > 0 Then .Body.Walk(.Heading).Started = 1
                .Arma.WeaponWalk(.Heading).Started = 1
                .Escudo.ShieldWalk(.Heading).Started = 1
                
                'Char moved
                moved = True
                
                'Check if we already got there
                If (Sgn(.scrollDirectionY) = 1 And .MoveOffsetY >= 0) Or (Sgn(.scrollDirectionY) = -1 And .MoveOffsetY <= 0) Then
                    .MoveOffsetY = 0
                    .scrollDirectionY = 0
                End If
            End If
        End If
        
        'If done moving stop animation
        If Not moved Then
            'Stop animations
            .Body.Walk(.Heading).Started = 0
            .Body.Walk(.Heading).FrameCounter = 1
            
            .Arma.WeaponWalk(.Heading).Started = 0
            .Arma.WeaponWalk(.Heading).FrameCounter = 1
            
            .Escudo.ShieldWalk(.Heading).Started = 0
            .Escudo.ShieldWalk(.Heading).FrameCounter = 1
            
            .Moving = False
        End If
        
        PixelOffsetX = PixelOffsetX + .MoveOffsetX
        PixelOffsetY = PixelOffsetY + .MoveOffsetY
        
        If .Head.Head(.Heading).grhindex Then
            If Not .invisible Then
                'Draw Body

                If .Body.Walk(.Heading).grhindex Then Call Draw_Grh(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, , x, y)
                                
                If .Head.Head(.Heading).grhindex Then
                    Call Draw_Grh(.Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.y, 1, 0, , x, y)

                    If .Casco.Head(.Heading).grhindex Then Call Draw_Grh(.Casco.Head(.Heading), PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.y, 1, 0, , x, y)

                    If .Arma.WeaponWalk(.Heading).grhindex Then Call Draw_Grh(.Arma.WeaponWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, , x, y)

                    If .Escudo.ShieldWalk(.Heading).grhindex Then Call Draw_Grh(.Escudo.ShieldWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, , x, y)
                End If
                            
                'Draw name over head
                If Nombres Then
                    If Len(.Nombre) > 0 Then
                        Pos = InStr(.Nombre, "<")

                        If Pos = 0 Then Pos = Len(.Nombre) + 2
                            
                        If .priv = 0 Then
                            If .Criminal Then
                                color = D3DColorXRGB(ColoresPJ(50).r, ColoresPJ(50).g, ColoresPJ(50).B)
                            Else
                                color = D3DColorXRGB(ColoresPJ(49).r, ColoresPJ(49).g, ColoresPJ(49).B)
                            End If

                        Else
                            color = D3DColorXRGB(ColoresPJ(.priv).r, ColoresPJ(.priv).g, ColoresPJ(.priv).B)
                        End If
                            
                        'Nick
                        line = Left$(.Nombre, Pos - 2)
                        Call DibujarTexto(line, PixelOffsetY + 30, PixelOffsetX - 10, 1, color)
                            
                        'Clan
                        line = mid$(.Nombre, Pos)
                        Call DibujarTexto(line, PixelOffsetY + 45, PixelOffsetX - 22, 1, color)
                    End If
                End If
            End If

            'End If
        Else

            'Draw Body
            If .Body.Walk(.Heading).grhindex Then Call Draw_Grh(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1)
        End If

        ''Update dialogs
        Call Dialogos.UpdateDialogPos(PixelOffsetX + .Body.HeadOffset.x - 155, PixelOffsetY + .Body.HeadOffset.y, charindex)
        
        'Draw FX
        If .FxIndex <> 0 Then
            Call Draw_Grh(.fX, PixelOffsetX + FxData(.FxIndex).OffsetX, PixelOffsetY + FxData(.FxIndex).OffsetY, 1, 1, True)

            If .fX.Started = 0 Then .FxIndex = 0
        End If

    End With

End Sub

Public Sub Start()

    Dim ulttick        As Long, esttick As Long

    Dim timers(1 To 2) As Integer

    Dim LoopC          As Integer

    day_r_old = luz_dia(Hour(Time)).r
    day_g_old = luz_dia(Hour(Time)).g
    day_b_old = luz_dia(Hour(Time)).B
    base_light = ARGB(day_r_old, day_g_old, day_b_old, 255)
    
    DoEvents
    vertList(0).x = 0: vertList(0).y = 0
    vertList(1).x = 800: vertList(1).y = 0
    vertList(2).x = 0: vertList(2).y = 600
    vertList(3).x = 800: vertList(3).y = 600
    vertList(0).rhw = 1
    vertList(1).rhw = 1
    vertList(2).rhw = 1
    vertList(3).rhw = 1
    vertList(0).color = D3DColorXRGB(255, 255, 255)
    vertList(1).color = D3DColorXRGB(255, 255, 255)
    vertList(2).color = D3DColorXRGB(255, 255, 255)
    vertList(3).color = D3DColorXRGB(255, 255, 255)
    vertList(0).tu = 0
    vertList(0).tv = 0
    vertList(1).tu = 1
    vertList(1).tv = 0
    vertList(2).tu = 0
    vertList(2).tv = 1
    vertList(3).tu = 1
    vertList(3).tv = 1

    On Error Resume Next ' putos timers!!!!!

    Do While prgRun

        If frmMain.WindowState <> vbMinimized And frmMain.Visible = True Then
            CheckKeys
            engine.Render
            engine.DrawInv
        Else
            Sleep 10&
        End If

        DoEvents

        esttick = GetTickCount

        For LoopC = 1 To UBound(timers)
            timers(LoopC) = timers(LoopC) + (esttick - ulttick)

            'Timer de trabajo
            If timers(1) >= tUs Then
                timers(1) = 0
                NoPuedeUsar = False
            End If

            'timer de attaque (77)
            If timers(2) >= tAt Then
                timers(2) = 0
                UserCanAttack = 1
                UserPuedeRefrescar = True
            End If

        Next LoopC

        ulttick = GetTickCount

        DoEvents
    Loop

    engine.Engine_Deinit
    
    frmCargando.Show

    'Destruimos los objetos públicos creados
    Set SurfaceDB = Nothing
    Set Dialogos = Nothing
    Set DialogosClanes = Nothing
    Set Audio = Nothing
    Set Inventario = Nothing
    
    Call UnloadAllForms
    
    End

End Sub

Public Sub SetCharacterFx(ByVal charindex As Integer, _
                          ByVal fX As Integer, _
                          ByVal Loops As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 12/03/04
    'Sets an FX to the character.
    '***************************************************
    With charlist(charindex)
        .FxIndex = fX
        
        If .FxIndex > 0 Then
            Call InitGrh(.fX, FxData(fX).Animacion)
        
            .fX.Loops = Loops
        End If

    End With

End Sub

Public Sub Char_Move_by_Head(ByVal charindex As Integer, ByVal nHeading As E_Heading)

    '*****************************************************************
    'Starts the movement of a character in nHeading direction
    '*****************************************************************
    Dim addx As Integer

    Dim addy As Integer

    Dim x    As Integer

    Dim y    As Integer

    Dim nX   As Integer

    Dim nY   As Integer
    
    With charlist(charindex)
        x = .Pos.x
        y = .Pos.y
        
        'Figure out which way to move
        Select Case nHeading

            Case E_Heading.NORTH
                addy = -1
        
            Case E_Heading.EAST
                addx = 1
        
            Case E_Heading.SOUTH
                addy = 1
            
            Case E_Heading.WEST
                addx = -1
        End Select
        
        nX = x + addx
        nY = y + addy
        
        MapData(nX, nY).charindex = charindex
        .Pos.x = nX
        .Pos.y = nY
        MapData(x, y).charindex = 0
        
        .MoveOffsetX = -1 * (32 * addx)
        .MoveOffsetY = -1 * (32 * addy)
        
        .Moving = 1
        .Heading = nHeading
        
        .scrollDirectionX = addx
        .scrollDirectionY = addy
    End With
    
    If UserEstado <> 1 Then Call DoPasosFx(charindex)
    
    'areas viejos
    If (nY < MinLimiteY) Or (nY > MaxLimiteY) Or (nX < MinLimiteX) Or (nX > MaxLimiteX) Then
        Call EraseChar(charindex)
    End If

End Sub

Public Sub Char_Move_by_Pos(ByVal charindex As Integer, _
                            ByVal nX As Integer, _
                            ByVal nY As Integer)

    On Error Resume Next

    Dim x        As Integer

    Dim y        As Integer

    Dim addx     As Integer

    Dim addy     As Integer

    Dim nHeading As E_Heading
    
    With charlist(charindex)
        x = .Pos.x
        y = .Pos.y
        
        MapData(x, y).charindex = 0
        
        addx = nX - x
        addy = nY - y
        
        If Sgn(addx) = 1 Then
            nHeading = E_Heading.EAST
        End If
        
        If Sgn(addx) = -1 Then
            nHeading = E_Heading.WEST
        End If
        
        If Sgn(addy) = -1 Then
            nHeading = E_Heading.NORTH
        End If
        
        If Sgn(addy) = 1 Then
            nHeading = E_Heading.SOUTH
        End If
        
        MapData(nX, nY).charindex = charindex
        
        .Pos.x = nX
        .Pos.y = nY
        
        .MoveOffsetX = -1 * (32 * addx)
        .MoveOffsetY = -1 * (32 * addy)
        
        .Moving = 1
        .Heading = nHeading
        
        .scrollDirectionX = Sgn(addx)
        .scrollDirectionY = Sgn(addy)
        
        'parche para que no medite cuando camina
        If .FxIndex = FxMeditar.FXMEDITARCHICO Or .FxIndex = FxMeditar.FXMEDITARGRANDE Or .FxIndex = FxMeditar.FXMEDITARMEDIANO Or .FxIndex = FxMeditar.FXMEDITARNIVEL50 Or .FxIndex = FxMeditar.FXMEDITARXGRANDE Then
            .FxIndex = 0
        End If

    End With
    
    If Not EstaPCarea(charindex) Then
        Call Dialogos.RemoveDialog(charindex)
        'Call Hits.RemoveHit(CharIndex)
    End If

    If (nY < MinLimiteY) Or (nY > MaxLimiteY) Or (nX < MinLimiteX) Or (nX > MaxLimiteX) Then
        Call EraseChar(charindex)
    End If

End Sub

Private Function EstaPCarea(ByVal charindex As Integer) As Boolean

    With charlist(charindex).Pos
        EstaPCarea = .x > UserPos.x - MinXBorder And .x < UserPos.x + MinXBorder And .y > UserPos.y - MinYBorder And .y < UserPos.y + MinYBorder
    End With

End Function

Private Function CreateColorVal(A As Integer, _
                                r As Integer, _
                                g As Integer, _
                                B As Integer) As D3DCOLORVALUE
    CreateColorVal.A = A
    CreateColorVal.r = r
    CreateColorVal.g = g
    CreateColorVal.B = B
End Function
Private Function Engine_FToDW(f As Single) As Long

    ' single > long
    Dim buf As D3DXBuffer

    Set buf = D3DX.CreateBuffer(4)
    D3DX.BufferSetData buf, 0, 4, 1, f
    D3DX.BufferGetData buf, 0, 4, 1, Engine_FToDW
End Function

Public Sub RemoveCharAparence(ByVal charindex As Integer, _
                              ByVal Muerto As Integer, _
                              ByVal Body As Integer, _
                              ByVal Head As Integer, _
                              ByVal Heading As E_Heading, _
                              ByVal tempint As Integer, _
                              ByVal tempweap As Integer, _
                              ByVal tempesc As Integer, _
                              ByVal fX As Integer, _
                              ByVal Loops As Integer)
                        
    charindex = charindex
    charlist(charindex).Muerto = Muerto = 500
    charlist(charindex).Body = BodyData(Body)
    charlist(charindex).Head = HeadData(Head)
    charlist(charindex).Heading = Heading
    charlist(charindex).FxIndex = fX
    charlist(charindex).fX.Loops = Loops
    tempint = tempint

    If tempint <> 0 Then charlist(charindex).Arma = WeaponAnimData(tempint)
    tempint = tempweap

    If tempint <> 0 Then charlist(charindex).Escudo = ShieldAnimData(tempint)
    tempint = tempesc

    If tempint <> 0 Then charlist(charindex).Casco = CascoAnimData(tempint)
    Call SetCharacterFx(charindex, charlist(charindex).FxIndex, charlist(charindex).fX.Loops)
            
    Call RefreshAllChars
End Sub

Public Sub Draw_FilledBox(ByVal x As Integer, _
                          ByVal y As Integer, _
                          ByVal width As Integer, _
                          ByVal height As Integer, _
                          color As Long, _
                          outlinecolor As Long)

    Static box_rect     As RECT

    Static Outline      As RECT

    Static rgb_list(3)  As Long

    Static rgb_list2(3) As Long

    Static Vertex(3)    As TLVERTEX

    Static Vertex2(3)   As TLVERTEX
    
    rgb_list(0) = color
    rgb_list(1) = color
    rgb_list(2) = color
    rgb_list(3) = color
    
    rgb_list2(0) = outlinecolor
    rgb_list2(1) = outlinecolor
    rgb_list2(2) = outlinecolor
    rgb_list2(3) = outlinecolor
    
    With box_rect
        .bottom = y + height
        .Left = x
        .Right = x + width
        .Top = y
    End With
    
    With Outline
        .bottom = y + height + 2
        .Left = x - 2
        .Right = x + width + 2
        .Top = y - 2
    End With
    
    Geometry_Create_Box Vertex2(), Outline, Outline, rgb_list2(), 0, 0
    Geometry_Create_Box Vertex(), box_rect, box_rect, rgb_list(), 0, 0
    
    D3DDevice.SetTexture 0, Nothing
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Vertex2(0), Len(Vertex2(0))
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Vertex(0), Len(Vertex(0))
End Sub

Public Sub DrawInv()

    Static re As RECT

    re.Left = 0
    re.Top = 0
    re.bottom = 160
    re.Right = 160
    
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 0, 0
    'D3DDevice.BeginScene
    Inventario.UpdateInventory
    'D3DDevice.EndScene
    D3DDevice.Present InvRect, ByVal 0, frmMain.picInv.hWnd, ByVal 0
End Sub

Public Sub GrhRenderToHdc(ByVal Grh_Index As Long, _
                          desthDC As Long, _
                          ByVal screen_x As Integer, _
                          ByVal screen_y As Integer, _
                          Optional transparent As Boolean = False)
                          On Error Resume Next
 
    Dim file_path  As String

    Dim src_x      As Integer

    Dim src_y      As Integer

    Dim src_width  As Integer

    Dim src_height As Integer

    Dim hdcsrc     As Long

    Dim MaskDC     As Long

    Dim PrevObj    As Long

    Dim PrevObj2   As Long
 
    If Grh_Index <= 0 Then Exit Sub
 
    'If it's animated switch grh_index to first frame
    If GrhData(Grh_Index).NumFrames <> 1 Then
        Grh_Index = GrhData(Grh_Index).Frames(1)
    End If
 
    file_path = modCompression.Extract_File(graphics, App.path & "\GRAFICOS\", GrhData(Grh_Index).FileNum & ".bmp", App.path & "\WAV\")
       
    src_x = GrhData(Grh_Index).sX
    src_y = GrhData(Grh_Index).sY
    src_width = GrhData(Grh_Index).pixelWidth
    src_height = GrhData(Grh_Index).pixelHeight
           
    hdcsrc = CreateCompatibleDC(desthDC)
    PrevObj = SelectObject(hdcsrc, LoadPicture(file_path))
       
    If transparent = False Then
        BitBlt desthDC, screen_x, screen_y, src_width, src_height, hdcsrc, src_x, src_y, vbSrcCopy
    Else
        MaskDC = CreateCompatibleDC(desthDC)
           
        PrevObj2 = SelectObject(MaskDC, LoadPicture(file_path))
           
        Grh_Create_Mask hdcsrc, MaskDC, src_x, src_y, src_width, src_height
           
        'Render tranparently
        BitBlt desthDC, screen_x, screen_y, src_width, src_height, MaskDC, src_x, src_y, vbSrcAnd
        BitBlt desthDC, screen_x, screen_y, src_width, src_height, hdcsrc, src_x, src_y, vbSrcPaint
           
        Call DeleteObject(SelectObject(MaskDC, PrevObj2))
           
        DeleteDC MaskDC
    End If
       
    Call DeleteObject(SelectObject(hdcsrc, PrevObj))
    DeleteDC hdcsrc
 
    Exit Sub

End Sub
 
Private Sub Grh_Create_Mask(ByRef hdcsrc As Long, _
                            ByRef MaskDC As Long, _
                            ByVal src_x As Integer, _
                            ByVal src_y As Integer, _
                            ByVal src_width As Integer, _
                            ByVal src_height As Integer)

    Dim x          As Integer

    Dim y          As Integer

    Dim TransColor As Long
   
    'ColorKey = hex(COLOR_KEY)
   
    'Check if it has an alpha component
    'If Len(ColorKey) > 6 Then
    'get rid of alpha
    '    ColorKey = "&H" & Right$(ColorKey, 6)
    'End If
    'piluex prueba
    'TransColor = Val(ColorKey)
    TransColor = &H0
 
    'Make it a mask (set background to black and foreground to white)
    'And set the sprite's background white
    For y = src_y To src_height + src_y
        For x = src_x To src_width + src_x

            If GetPixel(hdcsrc, x, y) = TransColor Then
                SetPixel MaskDC, x, y, vbWhite
                SetPixel hdcsrc, x, y, vbBlack
            Else
                SetPixel MaskDC, x, y, vbBlack
            End If

        Next x
    Next y

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''[PARTICULAS]''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Particle_Group_Create(ByVal map_x As Integer, ByVal map_y As Integer, ByRef grh_index_list() As Long, ByRef rgb_list() As Long, _
                                        Optional ByVal particle_count As Long = 20, Optional ByVal stream_type As Long = 1, _
                                        Optional ByVal alpha_blend As Boolean, Optional ByVal alive_counter As Long = -1, _
                                        Optional ByVal frame_speed As Single = 0.5, Optional ByVal id As Long, _
                                        Optional ByVal X1 As Integer, Optional ByVal Y1 As Integer, Optional ByVal angle As Integer, _
                                        Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                                        Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                                        Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                                        Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                                        Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                                        Optional bounce_strength As Long, Optional ByVal X2 As Integer, Optional ByVal Y2 As Integer, _
                                        Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                                        Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                                        Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional grh_resize As Boolean, _
                                        Optional grh_resizex As Integer, Optional grh_resizey As Integer)
'**************************************************************
'Author: Aaron Perkins
'Modified by: Ryan Cain (Onezero)
'Last Modify Date: 5/14/2003
'Returns the particle_group_index if successful, else 0
'Modified by Juan Martín Sotuyo Dodero
'Modified by Augusto José Rando
'**************************************************************
   
    If (map_x <> -1) And (map_y <> -1) Then
        If Map_Particle_Group_Get(map_x, map_y) = 0 Then
            Particle_Group_Create = Particle_Group_Next_Open
            Particle_Group_Make Particle_Group_Create, map_x, map_y, particle_count, stream_type, grh_index_list(), rgb_list(), alpha_blend, alive_counter, frame_speed, id, X1, Y1, angle, vecx1, vecx2, vecy1, vecy2, life1, life2, fric, spin_speedL, gravity, grav_strength, bounce_strength, X2, Y2, XMove, move_x1, move_x2, move_y1, move_y2, YMove, spin_speedH, spin, grh_resize, grh_resizex, grh_resizey
        End If
    Else
        Particle_Group_Create = Particle_Group_Next_Open
        Particle_Group_Make Particle_Group_Create, map_x, map_y, particle_count, stream_type, grh_index_list(), rgb_list(), alpha_blend, alive_counter, frame_speed, id, X1, Y1, angle, vecx1, vecx2, vecy1, vecy2, life1, life2, fric, spin_speedL, gravity, grav_strength, bounce_strength, X2, Y2, XMove, move_x1, move_x2, move_y1, move_y2, YMove, spin_speedH, spin, grh_resize, grh_resizex, grh_resizey
    End If
 
End Function
 
Public Function Particle_Group_Remove(ByVal particle_group_index As Long) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'*****************************************************************
    'Make sure it's a legal index
    If Particle_Group_Check(particle_group_index) Then
        Particle_Group_Destroy particle_group_index
        Particle_Group_Remove = True
    End If
End Function
 
Public Function Particle_Group_Remove_All() As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'*****************************************************************
    Dim Index As Long
   
    For Index = 1 To particle_group_last
        'Make sure it's a legal index
        If Particle_Group_Check(Index) Then
            Particle_Group_Destroy Index
        End If
    Next Index
   
    Particle_Group_Remove_All = True
End Function
 
Public Function Particle_Group_Find(ByVal id As Long) As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'Find the index related to the handle
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim LoopC As Long
   
    LoopC = 1
    Do Until particle_group_list(LoopC).id = id
        If LoopC = particle_group_last Then
            Particle_Group_Find = 0
            Exit Function
        End If
        LoopC = LoopC + 1
    Loop
   
    Particle_Group_Find = LoopC
Exit Function
ErrorHandler:
    Particle_Group_Find = 0
End Function
 
Private Sub Particle_Group_Destroy(ByVal particle_group_index As Long)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim temp As particle_group
    Dim i As Integer
   
    If particle_group_list(particle_group_index).map_x > 0 And particle_group_list(particle_group_index).map_y > 0 Then
        MapData(particle_group_list(particle_group_index).map_x, particle_group_list(particle_group_index).map_y).particle_group_index = 0
    ElseIf particle_group_list(particle_group_index).char_index Then
        If Char_Check(particle_group_list(particle_group_index).char_index) Then
            'For I = 1 To charlist(particle_group_list(particle_group_index).char_index).particle_count
            '    If charlist(particle_group_list(particle_group_index).char_index).particle_group(I) = particle_group_index Then
            '        charlist(particle_group_list(particle_group_index).char_index).particle_group(I) = 0
            '
            '        Exit For
            '    End If
            'Next I
        End If
    End If
   
    particle_group_list(particle_group_index) = temp
   
    'Update array size
    If particle_group_index = particle_group_last Then
        Do Until particle_group_list(particle_group_last).active
            particle_group_last = particle_group_last - 1
            If particle_group_last = 0 Then
                particle_group_count = 0
                Exit Sub
            End If
        Loop
        ReDim Preserve particle_group_list(1 To particle_group_last)
    End If
    particle_group_count = particle_group_count - 1
End Sub
 
 
Private Sub Particle_Group_Make(ByVal particle_group_index As Long, ByVal map_x As Integer, ByVal map_y As Integer, _
                                ByVal particle_count As Long, ByVal stream_type As Long, ByRef grh_index_list() As Long, ByRef rgb_list() As Long, _
                                Optional ByVal alpha_blend As Boolean, Optional ByVal alive_counter As Long = -1, _
                                Optional ByVal frame_speed As Single = 0.5, Optional ByVal id As Long, _
                                Optional ByVal X1 As Integer, Optional ByVal Y1 As Integer, Optional ByVal angle As Integer, _
                                Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                                Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                                Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                                Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                                Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                                Optional bounce_strength As Long, Optional ByVal X2 As Integer, Optional ByVal Y2 As Integer, _
                                Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                                Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                                Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional grh_resize As Boolean, _
                                Optional grh_resizex As Integer, Optional grh_resizey As Integer)
                               
'*****************************************************************
'Author: Aaron Perkins
'Modified by: Ryan Cain (Onezero)
'Last Modify Date: 5/15/2003
'Makes a new particle effect
'Modified by Juan Martín Sotuyo Dodero
'*****************************************************************
    'Update array size
    If particle_group_index > particle_group_last Then
        particle_group_last = particle_group_index
        ReDim Preserve particle_group_list(1 To particle_group_last)
    End If
    particle_group_count = particle_group_count + 1
   
    'Make active
    particle_group_list(particle_group_index).active = True
   
    'Map pos
    If (map_x <> -1) And (map_y <> -1) Then
        particle_group_list(particle_group_index).map_x = map_x
        particle_group_list(particle_group_index).map_y = map_y
    End If
   
    'Grh list
    ReDim particle_group_list(particle_group_index).grh_index_list(1 To UBound(grh_index_list))
    particle_group_list(particle_group_index).grh_index_list() = grh_index_list()
    particle_group_list(particle_group_index).grh_index_count = UBound(grh_index_list)
   
    'Sets alive vars
    If alive_counter = -1 Then
        particle_group_list(particle_group_index).alive_counter = -1
        particle_group_list(particle_group_index).never_die = True
    Else
        particle_group_list(particle_group_index).alive_counter = alive_counter
        particle_group_list(particle_group_index).never_die = False
    End If
   
    'alpha blending
    particle_group_list(particle_group_index).alpha_blend = alpha_blend
   
    'stream type
    particle_group_list(particle_group_index).stream_type = stream_type
   
    'speed
    particle_group_list(particle_group_index).frame_speed = frame_speed
   
    particle_group_list(particle_group_index).X1 = X1
    particle_group_list(particle_group_index).Y1 = Y1
    particle_group_list(particle_group_index).X2 = X2
    particle_group_list(particle_group_index).Y2 = Y2
    particle_group_list(particle_group_index).angle = angle
    particle_group_list(particle_group_index).vecx1 = vecx1
    particle_group_list(particle_group_index).vecx2 = vecx2
    particle_group_list(particle_group_index).vecy1 = vecy1
    particle_group_list(particle_group_index).vecy2 = vecy2
    particle_group_list(particle_group_index).life1 = life1
    particle_group_list(particle_group_index).life2 = life2
    particle_group_list(particle_group_index).fric = fric
    particle_group_list(particle_group_index).spin = spin
    particle_group_list(particle_group_index).spin_speedL = spin_speedL
    particle_group_list(particle_group_index).spin_speedH = spin_speedH
    particle_group_list(particle_group_index).gravity = gravity
    particle_group_list(particle_group_index).grav_strength = grav_strength
    particle_group_list(particle_group_index).bounce_strength = bounce_strength
    particle_group_list(particle_group_index).XMove = XMove
    particle_group_list(particle_group_index).YMove = YMove
    particle_group_list(particle_group_index).move_x1 = move_x1
    particle_group_list(particle_group_index).move_x2 = move_x2
    particle_group_list(particle_group_index).move_y1 = move_y1
    particle_group_list(particle_group_index).move_y2 = move_y2
   
    particle_group_list(particle_group_index).rgb_list(0) = rgb_list(0)
    particle_group_list(particle_group_index).rgb_list(1) = rgb_list(1)
    particle_group_list(particle_group_index).rgb_list(2) = rgb_list(2)
    particle_group_list(particle_group_index).rgb_list(3) = rgb_list(3)
   
    particle_group_list(particle_group_index).grh_resize = grh_resize
    particle_group_list(particle_group_index).grh_resizex = grh_resizex
    particle_group_list(particle_group_index).grh_resizey = grh_resizey
   
    'handle
    particle_group_list(particle_group_index).id = id
   
    'create particle stream
    particle_group_list(particle_group_index).particle_count = particle_count
    ReDim particle_group_list(particle_group_index).particle_stream(1 To particle_count)
   
    'plot particle group on map
    If (map_x <> -1) And (map_y <> -1) Then
        MapData(map_x, map_y).particle_group_index = particle_group_index
    End If
   
End Sub
Public Function Particle_Type_Get(ByVal particle_index As Long) As Long
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'Last Modify Date: 8/27/2003
'Returns the stream type of a particle stream
'*****************************************************************
    If Particle_Group_Check(particle_index) Then
        Particle_Type_Get = particle_group_list(particle_index).stream_type
    End If
End Function
Private Sub Particle_Group_Render(ByVal particle_group_index As Long, ByVal screen_x As Integer, ByVal screen_y As Integer)
'*****************************************************************
'Author: Aaron Perkins
'Modified by: Ryan Cain (Onezero)
'Modified by: Juan Martín Sotuyo Dodero
'Last Modify Date: 5/15/2003
'Renders a particle stream at a paticular screen point
'*****************************************************************
    Dim LoopC As Long
    Dim temp_rgb(0 To 3) As Long
    Dim no_move As Boolean
   
    'Set colors
    If UserMinHP = 0 Then
        temp_rgb(0) = D3DColorARGB(particle_group_list(particle_group_index).alpha_blend, 255, 255, 255)
        temp_rgb(1) = D3DColorARGB(particle_group_list(particle_group_index).alpha_blend, 255, 255, 255)
        temp_rgb(2) = D3DColorARGB(particle_group_list(particle_group_index).alpha_blend, 255, 255, 255)
        temp_rgb(3) = D3DColorARGB(particle_group_list(particle_group_index).alpha_blend, 255, 255, 255)
    Else
        temp_rgb(0) = particle_group_list(particle_group_index).rgb_list(0)
        temp_rgb(1) = particle_group_list(particle_group_index).rgb_list(1)
        temp_rgb(2) = particle_group_list(particle_group_index).rgb_list(2)
        temp_rgb(3) = particle_group_list(particle_group_index).rgb_list(3)
    End If
   
    If particle_group_list(particle_group_index).alive_counter Then
   
        'See if it is time to move a particle
        particle_group_list(particle_group_index).frame_counter = particle_group_list(particle_group_index).frame_counter + timerTicksPerFrame
        If particle_group_list(particle_group_index).frame_counter > particle_group_list(particle_group_index).frame_speed Then
            particle_group_list(particle_group_index).frame_counter = 0
            no_move = False
        Else
            no_move = True
        End If
       
   
   
        'If it's still alive render all the particles inside
        For LoopC = 1 To particle_group_list(particle_group_index).particle_count
       
            'Render particle
            Particle_Render particle_group_list(particle_group_index).particle_stream(LoopC), _
                            screen_x, screen_y, _
                            particle_group_list(particle_group_index).grh_index_list(Round(RandomNumber(1, particle_group_list(particle_group_index).grh_index_count), 0)), _
                            temp_rgb(), _
                            particle_group_list(particle_group_index).alpha_blend, no_move, _
                            particle_group_list(particle_group_index).X1, particle_group_list(particle_group_index).Y1, particle_group_list(particle_group_index).angle, _
                            particle_group_list(particle_group_index).vecx1, particle_group_list(particle_group_index).vecx2, _
                            particle_group_list(particle_group_index).vecy1, particle_group_list(particle_group_index).vecy2, _
                            particle_group_list(particle_group_index).life1, particle_group_list(particle_group_index).life2, _
                            particle_group_list(particle_group_index).fric, particle_group_list(particle_group_index).spin_speedL, _
                            particle_group_list(particle_group_index).gravity, particle_group_list(particle_group_index).grav_strength, _
                            particle_group_list(particle_group_index).bounce_strength, particle_group_list(particle_group_index).X2, _
                            particle_group_list(particle_group_index).Y2, particle_group_list(particle_group_index).XMove, _
                            particle_group_list(particle_group_index).move_x1, particle_group_list(particle_group_index).move_x2, _
                            particle_group_list(particle_group_index).move_y1, particle_group_list(particle_group_index).move_y2, _
                            particle_group_list(particle_group_index).YMove, particle_group_list(particle_group_index).spin_speedH, _
                            particle_group_list(particle_group_index).spin, particle_group_list(particle_group_index).grh_resize, particle_group_list(particle_group_index).grh_resizex, particle_group_list(particle_group_index).grh_resizey
        Next LoopC
       
        If no_move = False Then
            'Update the group alive counter
            If particle_group_list(particle_group_index).never_die = False Then
                particle_group_list(particle_group_index).alive_counter = particle_group_list(particle_group_index).alive_counter - 1
            End If
        End If
   
    Else
        'If it's dead destroy it
        particle_group_list(particle_group_index).particle_count = particle_group_list(particle_group_index).particle_count - 1
        If particle_group_list(particle_group_index).particle_count <= 0 Then Particle_Group_Destroy particle_group_index
    End If
End Sub
 
 
Private Sub Particle_Render(ByRef temp_particle As Particle, ByVal screen_x As Integer, ByVal screen_y As Integer, _
                            ByVal Grh_Index As Long, ByRef rgb_list() As Long, _
                            Optional ByVal alpha_blend As Boolean, Optional ByVal no_move As Boolean, _
                            Optional ByVal X1 As Integer, Optional ByVal Y1 As Integer, Optional ByVal angle As Integer, _
                            Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                            Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                            Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                            Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                            Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                            Optional ByVal bounce_strength As Long, Optional ByVal X2 As Integer, Optional ByVal Y2 As Integer, _
                            Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                            Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                            Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional grh_resize As Boolean, _
                            Optional grh_resizex As Integer, Optional grh_resizey As Integer)
'**************************************************************
'Author: Aaron Perkins
'Modified by: Ryan Cain (Onezero)
'Modified by: Juan Martín Sotuyo Dodero
'Last Modify Date: 5/15/2003
'**************************************************************
    If no_move = False Then
        If temp_particle.alive_counter = 0 Then
            'Start new particle
            InitGrh temp_particle.Grh, Grh_Index, alpha_blend
            temp_particle.x = RandomNumber(X1, X2) - (base_tile_size / 2)
            temp_particle.y = RandomNumber(Y1, Y2) - (base_tile_size / 2)
            temp_particle.vector_x = RandomNumber(vecx1, vecx2)
            temp_particle.vector_y = RandomNumber(vecy1, vecy2)
            temp_particle.angle = angle
            temp_particle.alive_counter = RandomNumber(life1, life2)
            temp_particle.friction = fric
        Else
            'Continue old particle
            'Do gravity
            If gravity = True Then
                temp_particle.vector_y = temp_particle.vector_y + grav_strength
                If temp_particle.y > 0 Then
                    'bounce
                    temp_particle.vector_y = bounce_strength
                End If
            End If
            'Do rotation
            'If spin = True Then temp_particle.Grh.angle = temp_particle.Grh.angle + (RandomNumber(spin_speedL, spin_speedH) / 100)
            If temp_particle.angle >= 360 Then
                temp_particle.angle = 0
            End If
           
            If XMove = True Then temp_particle.vector_x = RandomNumber(move_x1, move_x2)
            If YMove = True Then temp_particle.vector_y = RandomNumber(move_y1, move_y2)
        End If
       
        'Add in vector
        temp_particle.x = temp_particle.x + (temp_particle.vector_x \ temp_particle.friction)
        temp_particle.y = temp_particle.y + (temp_particle.vector_y \ temp_particle.friction)
   
        'decrement counter
         temp_particle.alive_counter = temp_particle.alive_counter - 1
    End If
    'Draw it
    If grh_resize = True Then
        If temp_particle.Grh.grhindex Then
             Grh_Render_Advance temp_particle.Grh, temp_particle.x + screen_x, temp_particle.y + screen_y, grh_resizex, grh_resizey, rgb_list(), True, True, alpha_blend
            Exit Sub
        End If
    End If
 
    If temp_particle.Grh.grhindex Then
        Grh_Render temp_particle.Grh, temp_particle.x + screen_x, temp_particle.y + screen_y, rgb_list(), True, True, alpha_blend
    End If
End Sub
Private Sub Grh_Render(ByRef Grh As Grh, ByVal screen_x As Integer, ByVal screen_y As Integer, ByRef rgb_list() As Long, Optional ByVal h_centered As Boolean = True, Optional ByVal v_centered As Boolean = True, Optional ByVal alpha_blend As Boolean = False)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/28/2003
'Modified by Juan Martín Sotuyo Dodero
'Added centering
'**************************************************************
    Dim tile_width As Integer
    Dim tile_height As Integer
    Dim Grh_Index As Long
   
    If Grh.grhindex = 0 Then Exit Sub
       
    'Animation
    If Grh.Started Then
        Grh.FrameCounter = Grh.FrameCounter + (timerTicksPerFrame * Grh.Speed)
        If Grh.FrameCounter > GrhData(Grh.grhindex).NumFrames Then
            'If Grh.noloop Then
            '    Grh.FrameCounter = GrhData(Grh.GrhIndex).NumFrames
            'Else
                Grh.FrameCounter = 1
            'End If
        End If
    End If
 
        'particle_group_list(particle_group_index).frame_counter = particle_group_list(particle_group_index).frame_counter + timer_ticks_per_frame
        'If particle_group_list(particle_group_index).frame_counter > particle_group_list(particle_group_index).frame_speed Then
        '    particle_group_list(particle_group_index).frame_counter = 0
        '    no_move = False
        'Else
        '    no_move = True
        'End If
 
    'Figure out what frame to draw (always 1 if not animated)
    If Grh.FrameCounter = 0 Then Grh.FrameCounter = 1
    'If Not Grh_Check(Grh.grhindex) Then Exit Sub
    Grh_Index = GrhData(Grh.grhindex).Frames(Grh.FrameCounter)
    If Grh_Index <= 0 Then Exit Sub
    If GrhData(Grh_Index).FileNum = 0 Then Exit Sub
       
    'Modified by Augusto José Rando
    'Simplier function - according to basic ORE engine
    If h_centered Then
        If GrhData(Grh.grhindex).TileWidth <> 1 Then
            screen_x = screen_x - Int(GrhData(Grh.grhindex).TileWidth * (32 \ 2)) + 32 \ 2
        End If
    End If
   
    If v_centered Then
        If GrhData(Grh.grhindex).TileHeight <> 1 Then
            screen_y = screen_y - Int(GrhData(Grh.grhindex).TileHeight * 32) + 32
        End If
    End If
   
    'Draw it to device
    Device_Box_Textured_Render Grh_Index, _
        screen_x, screen_y, _
        GrhData(Grh_Index).pixelWidth, GrhData(Grh_Index).pixelHeight, _
        rgb_list(), _
        GrhData(Grh_Index).sX, GrhData(Grh_Index).sY, _
        alpha_blend ', _
        Grh.angle
 
End Sub
Private Sub Grh_Render_Advance(ByRef Grh As Grh, ByVal screen_x As Integer, ByVal screen_y As Integer, ByVal height As Integer, ByVal width As Integer, ByRef rgb_list() As Long, Optional ByVal h_center As Boolean, Optional ByVal v_center As Boolean, Optional ByVal alpha_blend As Boolean = False)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'Last Modify Date: 11/19/2003
'Similar to Grh_Render, but let´s you resize the Grh
'**************************************************************
    Dim tile_width As Integer
    Dim tile_height As Integer
    Dim Grh_Index As Long
   
    'Animation
    If Grh.Started Then
        Grh.FrameCounter = Grh.FrameCounter + (timerTicksPerFrame * Grh.Speed)
        If Grh.FrameCounter > GrhData(Grh.grhindex).NumFrames Then
            'If Grh.noloop Then
            '    Grh.FrameCounter = GrhData(Grh.GrhIndex).NumFrames
            'Else
                Grh.FrameCounter = 1
            'End If
        End If
    End If
   
    'Figure out what frame to draw (always 1 if not animated)
    If Grh.FrameCounter = 0 Then Grh.FrameCounter = 1
    Grh_Index = GrhData(Grh.grhindex).Frames(Grh.FrameCounter)
   
    'Center Grh over X, Y pos
    If GrhData(Grh.grhindex).TileWidth <> 1 Then
        screen_x = screen_x - Int(GrhData(Grh.grhindex).TileWidth * (base_tile_size \ 2)) + base_tile_size \ 2
    End If
   
    If GrhData(Grh.grhindex).TileHeight <> 1 Then
        screen_y = screen_y - Int(GrhData(Grh.grhindex).TileHeight * base_tile_size) + base_tile_size
    End If
   
    'Draw it to device
    Device_Box_Textured_Render_Advance Grh_Index, _
        screen_x, screen_y, _
        GrhData(Grh_Index).pixelWidth, GrhData(Grh_Index).pixelHeight, _
        rgb_list, _
        GrhData(Grh_Index).sX, GrhData(Grh_Index).sY, _
        width, height, alpha_blend ', grh.angle
End Sub
 
'Sub CARGARMAP()
'Particle_Group_Remove_All
'End Sub
Private Function Particle_Group_Next_Open() As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'*****************************************************************
On Error GoTo ErrorHandler:
    Dim LoopC As Long
   
    LoopC = 1
    Do Until particle_group_list(LoopC).active = False
        If LoopC = particle_group_last Then
            Particle_Group_Next_Open = particle_group_last + 1
            Exit Function
        End If
        LoopC = LoopC + 1
    Loop
   
    Particle_Group_Next_Open = LoopC
Exit Function
ErrorHandler:
    Particle_Group_Next_Open = 1
End Function
 
Private Function Particle_Group_Check(ByVal particle_group_index As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'**************************************************************
    'check index
    If particle_group_index > 0 And particle_group_index <= particle_group_last Then
        If particle_group_list(particle_group_index).active Then
            Particle_Group_Check = True
        End If
    End If
End Function

Public Function General_Particle_Create(ByVal ParticulaInd As Long, ByVal x As Integer, ByVal y As Integer, Optional ByVal particle_life As Long = 0) As Long
    
Dim rgb_list(0 To 3) As Long
rgb_list(0) = RGB(StreamData(ParticulaInd).colortint(0).r, StreamData(ParticulaInd).colortint(0).g, StreamData(ParticulaInd).colortint(0).B)
rgb_list(1) = RGB(StreamData(ParticulaInd).colortint(1).r, StreamData(ParticulaInd).colortint(1).g, StreamData(ParticulaInd).colortint(1).B)
rgb_list(2) = RGB(StreamData(ParticulaInd).colortint(2).r, StreamData(ParticulaInd).colortint(2).g, StreamData(ParticulaInd).colortint(2).B)
rgb_list(3) = RGB(StreamData(ParticulaInd).colortint(3).r, StreamData(ParticulaInd).colortint(3).g, StreamData(ParticulaInd).colortint(3).B)

General_Particle_Create = engine.Particle_Group_Create(x, y, StreamData(ParticulaInd).grh_list, rgb_list(), StreamData(ParticulaInd).NumOfParticles, ParticulaInd, _
    StreamData(ParticulaInd).AlphaBlend, IIf(particle_life = 0, StreamData(ParticulaInd).life_counter, particle_life), StreamData(ParticulaInd).Speed, , StreamData(ParticulaInd).X1, StreamData(ParticulaInd).Y1, StreamData(ParticulaInd).angle, _
    StreamData(ParticulaInd).vecx1, StreamData(ParticulaInd).vecx2, StreamData(ParticulaInd).vecy1, StreamData(ParticulaInd).vecy2, _
    StreamData(ParticulaInd).life1, StreamData(ParticulaInd).life2, StreamData(ParticulaInd).friction, StreamData(ParticulaInd).spin_speedL, _
    StreamData(ParticulaInd).gravity, StreamData(ParticulaInd).grav_strength, StreamData(ParticulaInd).bounce_strength, StreamData(ParticulaInd).X2, _
    StreamData(ParticulaInd).Y2, StreamData(ParticulaInd).XMove, StreamData(ParticulaInd).move_x1, StreamData(ParticulaInd).move_x2, StreamData(ParticulaInd).move_y1, _
    StreamData(ParticulaInd).move_y2, StreamData(ParticulaInd).YMove, StreamData(ParticulaInd).spin_speedH, StreamData(ParticulaInd).spin, StreamData(ParticulaInd).grh_resize, StreamData(ParticulaInd).grh_resizex, StreamData(ParticulaInd).grh_resizey)

End Function
Public Function General_Char_Particle_Create(ByVal ParticulaInd As Long, ByVal char_index As Integer, Optional ByVal particle_life As Long = 0) As Long

Dim rgb_list(0 To 3) As Long
rgb_list(0) = RGB(StreamData(ParticulaInd).colortint(0).r, StreamData(ParticulaInd).colortint(0).g, StreamData(ParticulaInd).colortint(0).B)
rgb_list(1) = RGB(StreamData(ParticulaInd).colortint(1).r, StreamData(ParticulaInd).colortint(1).g, StreamData(ParticulaInd).colortint(1).B)
rgb_list(2) = RGB(StreamData(ParticulaInd).colortint(2).r, StreamData(ParticulaInd).colortint(2).g, StreamData(ParticulaInd).colortint(2).B)
rgb_list(3) = RGB(StreamData(ParticulaInd).colortint(3).r, StreamData(ParticulaInd).colortint(3).g, StreamData(ParticulaInd).colortint(3).B)

'General_Char_Particle_Create = engine.Char_Particle_Group_Create(char_index, StreamData(ParticulaInd).grh_list, rgb_list(), StreamData(ParticulaInd).NumOfParticles, ParticulaInd, _
'    StreamData(ParticulaInd).AlphaBlend, IIf(particle_life = 0, StreamData(ParticulaInd).life_counter, particle_life), StreamData(ParticulaInd).Speed, , StreamData(ParticulaInd).x1, StreamData(ParticulaInd).y1, StreamData(ParticulaInd).angle, _
'    StreamData(ParticulaInd).vecx1, StreamData(ParticulaInd).vecx2, StreamData(ParticulaInd).vecy1, StreamData(ParticulaInd).vecy2, _
'    StreamData(ParticulaInd).life1, StreamData(ParticulaInd).life2, StreamData(ParticulaInd).friction, StreamData(ParticulaInd).spin_speedL, _
'    StreamData(ParticulaInd).gravity, StreamData(ParticulaInd).grav_strength, StreamData(ParticulaInd).bounce_strength, StreamData(ParticulaInd).x2, _
'    StreamData(ParticulaInd).y2, StreamData(ParticulaInd).XMove, StreamData(ParticulaInd).move_x1, StreamData(ParticulaInd).move_x2, StreamData(ParticulaInd).move_y1, _
'    StreamData(ParticulaInd).move_y2, StreamData(ParticulaInd).YMove, StreamData(ParticulaInd).spin_speedH, StreamData(ParticulaInd).spin, StreamData(ParticulaInd).grh_resize, StreamData(ParticulaInd).grh_resizex, StreamData(ParticulaInd).grh_resizey)

End Function

Public Function Map_Particle_Group_Get(ByVal map_x As Long, ByVal map_y As Long) As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/20/2003
'Checks to see if a tile position has a particle_group_index and return it
'*****************************************************************
    If Map_In_Bounds(map_x, map_y) Then
        Map_Particle_Group_Get = map_current.map_grid(map_x, map_y).particle_group_index
    Else
        Map_Particle_Group_Get = 0
    End If
End Function

Private Function Char_Check(ByVal char_index As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'**************************************************************
    'check char_index
    If char_index > 0 And char_index <= char_last Then
        If char_list(char_index).active Then
            Char_Check = True
        End If
    End If
End Function

Private Sub Device_Box_Textured_Render_Advance(ByVal grhindex As Long, ByVal dest_x As Integer, ByVal dest_y As Integer, ByVal src_width As Integer, _
                                            ByVal src_height As Integer, ByRef rgb_list() As Long, ByVal src_x As Integer, _
                                            ByVal src_y As Integer, ByVal dest_width As Integer, Optional ByVal dest_height As Integer, _
                                            Optional ByVal alpha_blend As Boolean, Optional ByVal angle As Single)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 5/15/2003
'Copies the Textures allowing resizing
'Modified by Juan Martín Sotuyo Dodero
'**************************************************************
    Static src_rect As RECT
    Static dest_rect As RECT
    Static temp_verts(3) As TLVERTEX
    Static d3dTextures As D3D8Textures
    Static light_value(0 To 3) As Long

    If grhindex = 0 Then Exit Sub
    Set d3dTextures.texture = SurfaceDB.GetTexture(GrhData(grhindex).FileNum, d3dTextures.texwidth, d3dTextures.texheight)
    
        light_value(0) = &HFFFFFFFF 'blind_color
        light_value(1) = &HFFFFFFFF 'blind_color
        light_value(2) = &HFFFFFFFF 'blind_color
        light_value(3) = &HFFFFFFFF 'blind_color
    
    With src_rect
        .bottom = src_y + src_height
        .Left = src_x
        .Right = src_x + src_width
        .Top = src_y
    End With
        
    'Set up the destination rectangle
    With dest_rect
        .bottom = dest_y + dest_height
        .Left = dest_x
        .Right = dest_x + dest_width
        .Top = dest_y
    End With
    
    'Set up the TempVerts(3) vertices
    Geometry_Create_Box temp_verts(), dest_rect, src_rect, light_value(), d3dTextures.texwidth, d3dTextures.texheight, angle
        
    'Set Textures
    D3DDevice.SetTexture 0, d3dTextures.texture
    
    If alpha_blend Then
       'Set Rendering for alphablending
        D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    End If
    
    'Draw the triangles that make up our square Textures
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, temp_verts(0), Len(temp_verts(0))
    
    If alpha_blend Then
        'Set Rendering for colokeying
        D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    End If

End Sub


Public Function Map_In_Bounds(ByVal map_x As Long, ByVal map_y As Long) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Checks to see if a tile position is in the maps bounds
'*****************************************************************
    If map_x < map_current.map_x_min Or map_x > map_current.map_x_max Or map_y < map_current.map_y_min Or map_y > map_current.map_y_max Then
        Map_In_Bounds = False
        Exit Function
    End If
   
    Map_In_Bounds = True
End Function

