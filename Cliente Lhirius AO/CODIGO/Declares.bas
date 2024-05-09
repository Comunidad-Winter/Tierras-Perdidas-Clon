Attribute VB_Name = "Mod_Declaraciones"

Option Explicit
Public InvRect As RECT
Public Movement_Speed As Single
Public Const INFINITE_LOOPS As Integer = -1
'Inventory Items, Kshameenk
Public MouseOK As Boolean
Public MouseItem As Integer
Public MouseRendOK As Boolean
Public OfMouse As Boolean
Public ButtonIN As Boolean
Public puedo As Boolean

Public NamePrivado As String
Public NameAmigo As String
Public ma(1) As Single
'Public Movement_Speed As Single
Public LuzMouse As Boolean
Public cuentapj As Integer

'Particle Groups
Public TotalStreams As Integer
Public StreamData() As Stream
 
'RGB Type
Public Type RGB
    r As Long
    g As Long
    B As Long
End Type
 
Public Type Stream
    grh_resize As Boolean
    grh_resizex As Integer
    grh_resizey As Integer
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
 Public nombrecuent As String
Public passcuent As String
'FISHAR.-
Public Type RGAOT
     Cupos As String
     Nivel As String
     TIPO As String
     AutoSum As Boolean
     m As Long
     x As Long
     Y As Long
End Type
Public Torneos As RGAOT
Public CartelInvisibilidad As Integer
Public CartelParalisis As Integer  ' Se entiende supongo
Public MapName As String
Type tMensajesSos
    TIPO As String
    Autor As String
    Contenido As String
End Type
 
Public MensajesSOS(1 To 500) As tMensajesSos
Public EsUsuario As String
Public MensajesNumber As Integer
Public TieneParaResponder As Boolean
Public SendingType As Byte
Public sndPrivateTo As String
Public mode As Boolean
Public temp_rgb(3) As Long
Public LuzGrh(3) As Long
Public AlphaY As Byte
 
Rem Mannakia Desvanecimientos
Public Type tDesv
    Act As Byte
    AlphaX As Byte
End Type
Public TechoDesv As tDesv
'juanjo
Public PJClickeado As String

Public rcvName As String
Public rcvHead As Integer
Public rcvBody As Integer
Public rcvShield As Integer
Public rcvWeapon As Integer
Public rcvCasco As Integer
Public rcvIndex As Integer
Public rcvCrimi As Boolean
Public rcvBaned As Integer
Public rcvLevel As Integer
Public rcvClase As String
Public rcvMuerto As Integer

Public PJSAmount As Integer
'/juanjo

'Objetos públicos
Public IsAttacking As Boolean
Public DialogosClanes As New clsGuildDlg
Public Dialogos As New cDialogos
Public Audio As New clsAudio
Public Light As New clsLight
Public Inventario As New clsGrapchicalInventory
Public SurfaceDB As clsSurfaceManDynDX8   'No va new porque es unainterfaz, el new se pone al decidir que clase de objeto es

#If SeguridadAlkon Then
Public md5 As New clsMD5
#End If

'For function GetTickCount
Public TickCount As Long

'Sonidos
Public Const SND_CLICK As String = "click.Wav"
Public Const SND_PASOS1 As String = "23.Wav"
Public Const SND_PASOS2 As String = "24.Wav"
Public Const SND_NAVEGANDO As String = "50.wav"
Public Const SND_OVER As String = "click2.Wav"
Public Const SND_DICE As String = "cupdice.Wav"
Public Const SND_LLUVIAINEND As String = "lluviainend.wav"
Public Const SND_LLUVIAOUTEND As String = "lluviaoutend.wav"

'Musica
Public Const MIdi_Inicio As Byte = 6

Public RawServersList As String
Public UsersOns As Integer

Public Type tColor
    r As Byte
    g As Byte
    B As Byte
End Type

Public ColoresPJ(0 To 50) As tColor


Public Type tServerInfo
    Ip As String
    Puerto As Integer
    desc As String
    PassRecPort As Integer
End Type

Public currentMidi As Long

Public ServersLst() As tServerInfo
Public ServersRecibidos As Boolean
Public MagMin As Integer
Public MagMax As Integer
Public MagMina As Integer
Public MagMaxa As Integer
Public MagMinb As Integer
Public MagMaxb As Integer
Public MagMinc As Integer
Public MagMaxc As Integer
Public MagMind As Integer
Public MagMaxd As Integer
Public CascMin As Integer
Public CascMax As Integer
Public EscuMin As Integer
Public EscuMax As Integer
Public ArmaMin As Integer
Public ArmaMax As Integer
Public ArmorMin As Integer
Public ArmorMax As Integer
Public HerrMin As Integer
Public HerrMax As Integer
Public CurServer As Integer

Public CreandoClan As Boolean
Public ClanName As String
Public Site As String

Public UserCiego As Boolean
Public UserEstupido As Boolean

Public NoRes As Boolean 'no cambiar la resolucion

Public RainBufferIndex As Long
Public FogataBufferIndex As Long

Public Const bCabeza = 1
Public Const bPiernaIzquierda = 2
Public Const bPiernaDerecha = 3
Public Const bBrazoDerecho = 4
Public Const bBrazoIzquierdo = 5
Public Const bTorso = 6

'Timers de GetTickCount
Public Const tAt = 2000
Public Const tUs = 600

Public Const PrimerBodyBarco = 84
Public Const UltimoBodyBarco = 87

Public NumEscudosAnims As Integer

Public ArmasHerrero(0 To 100) As Integer
Public ArmadurasHerrero(0 To 100) As Integer
Public ObjCarpintero(0 To 100) As Integer

Public Versiones(1 To 7) As Integer

Public UsaMacro As Boolean
Public CnTd As Byte
Public SecuenciaMacroHechizos As Byte



'[KEVIN]
Public Const MAX_BANCOINVENTORY_SLOTS = 40
Public UserBancoInventory(1 To MAX_BANCOINVENTORY_SLOTS) As Inventory
'[/KEVIN]


Public Tips() As String * 255
Public Const LoopAdEternum = 999

'Direcciones
Public Enum E_Heading
    NORTH = 1
    EAST = 2
    SOUTH = 3
    WEST = 4
End Enum

'Objetos
Public Const MAX_INVENTORY_OBJS = 10000
Public Const MAX_INVENTORY_SLOTS = 20
Public Const MAX_NPC_INVENTORY_SLOTS = 50
Public Const MAXHECHI = 35

Public Const MAXSKILLPOINTS = 100

Public Const FLAGORO = 777

Public Const FOgata = 1521

Public Enum Skills
     Suerte = 1
     Magia = 2
     Robar = 3
     Tacticas = 4
     Armas = 5
     Meditar = 6
     Apuñalar = 7
     Ocultarse = 8
     Supervivencia = 9
     Talar = 10
     Comerciar = 11
     Defensa = 12
     Pesca = 13
     Mineria = 14
     Carpinteria = 15
     Herreria = 16
     Liderazgo = 17 ' NOTA: Solia decir "Curacion"
     Domar = 18
     Proyectiles = 19
     Wresterling = 20
     Navegacion = 21
     Resistencia = 22
End Enum

Public Const FundirMetal As Integer = 88

'
' Mensajes
'
' MENSAJE_*  --> Mensajes de texto que se muestran en el cuadro de texto
'

Public Const MENSAJE_CRIATURA_FALLA_GOLPE As String = "La criatura fallo el golpe!!!"
Public Const MENSAJE_CRIATURA_MATADO As String = "La criatura te ha matado!!!"
Public Const MENSAJE_RECHAZO_ATAQUE_ESCUDO As String = "Has rechazado el ataque con el escudo!!!"
Public Const MENSAJE_USUARIO_RECHAZO_ATAQUE_ESCUDO  As String = "El usuario rechazo el ataque con su escudo!!!"
Public Const MENSAJE_FALLADO_GOLPE As String = "Has fallado el golpe!!!"
Public Const MENSAJE_SEGURO_ACTIVADO As String = ">>SEGURO ACTIVADO<<"
Public Const MENSAJE_SEGURO_DESACTIVADO As String = ">>SEGURO DESACTIVADO<<"
Public Const MENSAJE_PIERDE_NOBLEZA As String = "¡¡Has perdido puntaje de nobleza y ganado puntaje de criminalidad!! Si sigues ayudando a criminales te convertirás en uno de ellos y serás perseguido por las tropas de las ciudades."
Public Const MENSAJE_USAR_MEDITANDO As String = "¡Estás meditando! Debes dejar de meditar para usar objetos."

Public Const MENSAJE_GOLPE_CABEZA As String = "¡¡La criatura te ha pegado en la cabeza por "
Public Const MENSAJE_GOLPE_BRAZO_IZQ As String = "¡¡La criatura te ha pegado el brazo izquierdo por "
Public Const MENSAJE_GOLPE_BRAZO_DER As String = "¡¡La criatura te ha pegado el brazo derecho por "
Public Const MENSAJE_GOLPE_PIERNA_IZQ As String = "¡¡La criatura te ha pegado la pierna izquierda por "
Public Const MENSAJE_GOLPE_PIERNA_DER As String = "¡¡La criatura te ha pegado la pierna derecha por "
Public Const MENSAJE_GOLPE_TORSO  As String = "¡¡La criatura te ha pegado en el torso por "

' MENSAJE_[12]: Aparecen antes y despues del valor de los mensajes anteriores (MENSAJE_GOLPE_*)
Public Const MENSAJE_1 As String = "¡¡"
Public Const MENSAJE_2 As String = "!!"

Public Const MENSAJE_GOLPE_CRIATURA_1 As String = "¡¡Le has pegado a la criatura por "

Public Const MENSAJE_ATAQUE_FALLO As String = " te ataco y fallo!!"

Public Const MENSAJE_RECIVE_IMPACTO_CABEZA As String = " te ha pegado en la cabeza por "
Public Const MENSAJE_RECIVE_IMPACTO_BRAZO_IZQ As String = " te ha pegado el brazo izquierdo por "
Public Const MENSAJE_RECIVE_IMPACTO_BRAZO_DER As String = " te ha pegado el brazo derecho por "
Public Const MENSAJE_RECIVE_IMPACTO_PIERNA_IZQ As String = " te ha pegado la pierna izquierda por "
Public Const MENSAJE_RECIVE_IMPACTO_PIERNA_DER As String = " te ha pegado la pierna derecha por "
Public Const MENSAJE_RECIVE_IMPACTO_TORSO As String = " te ha pegado en el torso por "

Public Const MENSAJE_PRODUCE_IMPACTO_1 As String = "¡¡Le has pegado a "
Public Const MENSAJE_PRODUCE_IMPACTO_CABEZA As String = " en la cabeza por "
Public Const MENSAJE_PRODUCE_IMPACTO_BRAZO_IZQ As String = " en el brazo izquierdo por "
Public Const MENSAJE_PRODUCE_IMPACTO_BRAZO_DER As String = " en el brazo derecho por "
Public Const MENSAJE_PRODUCE_IMPACTO_PIERNA_IZQ As String = " en la pierna izquierda por "
Public Const MENSAJE_PRODUCE_IMPACTO_PIERNA_DER As String = " en la pierna derecha por "
Public Const MENSAJE_PRODUCE_IMPACTO_TORSO As String = " en el torso por "

Public Const MENSAJE_TRABAJO_MAGIA As String = "Haz click sobre el objetivo..."
Public Const MENSAJE_TRABAJO_PESCA As String = "Haz click sobre el sitio donde quieres pescar..."
Public Const MENSAJE_TRABAJO_ROBAR As String = "Haz click sobre la victima..."
Public Const MENSAJE_TRABAJO_TALAR As String = "Haz click sobre el árbol..."
Public Const MENSAJE_TRABAJO_MINERIA As String = "Haz click sobre el yacimiento..."
Public Const MENSAJE_TRABAJO_FUNDIRMETAL As String = "Haz click sobre la fragua..."
Public Const MENSAJE_TRABAJO_PROYECTILES As String = "Haz click sobre la victima..."

Public Const MENSAJE_ENTRAR_PARTY_1 As String = "Si deseas entrar en una party con "
Public Const MENSAJE_ENTRAR_PARTY_2 As String = ", escribe /entrarparty"

Public Const MENSAJE_NENE As String = "Cantidad de NPCs: "

'Inventario
Type Inventory
    OBJIndex As Integer
    Name As String
    grhindex As Integer
    '[Alejo]: tipo de datos ahora es Long
    Amount As Long
    '[/Alejo]
    Equipped As Byte
    Valor As Long
    OBJType As Integer
    def As Integer
    MaxHit As Integer
    MinHit As Integer
End Type

Type NpCinV
    OBJIndex As Integer
    Name As String
    grhindex As Integer
    Amount As Integer
    Valor As Long
    OBJType As Integer
    def As Integer
    MaxHit As Integer
    MinHit As Integer
    C1 As String
    C2 As String
    C3 As String
    C4 As String
    C5 As String
    C6 As String
    C7 As String
    
End Type

Type tReputacion 'Fama del usuario
    NobleRep As Long
    BurguesRep As Long
    PlebeRep As Long
    LadronesRep As Long
    BandidoRep As Long
    AsesinoRep As Long
    
    Promedio As Long
End Type

Type tEstadisticasUsu
    CiudadanosMatados As Long
    CriminalesMatados As Long
    UsuariosMatados As Long
    NpcsMatados As Long
    Clase As String
    PenaCarcel As Long
End Type

Public Nombres As Boolean

Public MixedKey As Long

'User status vars
Global OtroInventario(1 To MAX_INVENTORY_SLOTS) As Inventory

Public UserHechizos(1 To MAXHECHI) As Integer

Public NPCInventory(1 To MAX_NPC_INVENTORY_SLOTS) As NpCinV
Public NPCInvDim As Integer
Public UserMeditar As Boolean
Public UserName As String
Public UserPassword As String
Public UserMaxHP As Integer
Public UserMinHP As Integer
Public UserMaxMAN As Integer
Public UserMinMAN As Integer
Public UserMaxSTA As Integer
Public UserMinSTA As Integer
Public UserGLD As Long
Public UserHonor As Long
Public UserLvl As Integer
Public UserGLDBOV As Long
Public UserBOVItem As Long
Public UserPort As Integer
Public UserServerIP As String
Public UserCanAttack As Integer
Public UserEstado As Byte '0 = Vivo & 1 = Muerto
Public UserPasarNivel As Long
Public UserExp As Long
Public UserReputacion As tReputacion
Public UserEstadisticas As tEstadisticasUsu
Public UserDescansar As Boolean
Public tipf As String
Public PrimeraVez As Boolean
Public FPSFLAG As Boolean
Public pausa As Boolean
Public IScombate As Boolean
Public UserParalizado As Boolean
Public UserNavegando As Boolean
Public UserHogar As String
Public Type User_Info
    UserConsola As Byte
    UserDrag As Byte
    End Type
Public Usuario As User_Info

'<-------------------------NUEVO-------------------------->
Public Comerciando As Boolean
'<-------------------------NUEVO-------------------------->

Public UserClase As String
Public UserSexo As String
Public UserRaza As String
Public UserEmail As String


Public UserCounters As tUserCounters
Type tUserCounters
    ToNextFaction As Integer
    MVPs As Integer
    Wars As Integer
    CVC As Integer
    DuelWins As Integer
    DuelLoss As Integer
    Gold As Integer
    Silver As Integer
    Bronze As Integer
    Quest As Integer
    Warns As Integer
    KilledNPC As Integer
    Jerarquia As Integer
    Alineacion As Integer
End Type

Public Const NUMCIUDADES As Byte = 1
Public Const NUMSKILLS As Byte = 22
Public Const NUMATRIBUTOS As Byte = 5
Public Const NUMCLASES As Byte = 16
Public Const NUMRAZAS As Byte = 5

Public UserSkills(1 To NUMSKILLS) As Integer
Public SkillsNames(1 To NUMSKILLS) As String

Public UserAtributos(1 To NUMATRIBUTOS) As Integer
Public AtributosNames(1 To NUMATRIBUTOS) As String

Public Ciudades(1) As String
Public CityDesc(1) As String

Public ListaRazas(1 To NUMRAZAS) As String
Public ListaClases(1 To NUMCLASES) As String

Public Musica As Boolean
Public Sound As Boolean

Public SkillPoints As Integer
Public Alocados As Integer
Public flags() As Integer
Public Oscuridad As Integer
Public logged As Boolean
Public NoPuedeUsar As Boolean

'Barrin 30/9/03
Public UserPuedeRefrescar As Boolean

Public UsingSkill As Integer


Public MD5HushYo As String * 16

Public Enum E_MODO
    Normal = 1
    CrearNuevoPj = 2
    Dados = 3
    CrearAccount = 4
    LoginAccount = 5
    BorrarPj = 6
End Enum

Public EstadoLogin As E_MODO
   
Public Enum FxMeditar
    FXMEDITARCHICO = 4
    FXMEDITARMEDIANO = 5
    FXMEDITARGRANDE = 6
    FXMEDITARXGRANDE = 16
    FXMEDITARNIVEL50 = 23
End Enum


'Server stuff
Public stxtbuffer As String 'Holds temp raw data from server
Public stxtbuffercmsg As String 'Holds temp raw data from server
Public SendNewChar As Boolean 'Used during login
Public Connected As Boolean 'True when connected to server
Public UserMap As Integer

'String contants
Public Const ENDC As String * 1 = vbNullChar    'Endline character for talking with server
Public Const ENDL As String * 2 = vbCrLf        'Holds the Endline character for textboxes

'Control
Public prgRun As Boolean 'When true the program ends

Public IPdelServidor As String
Public PuertoDelServidor As String

'
'********** FUNCIONES API ***********
'

Public Declare Function GetTickCount Lib "kernel32" () As Long

'para escribir y leer variables
Public Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Public Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

'Invi con Alpha
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

'Teclado
Public Declare Function GetKeyState Lib "User32" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetAsyncKeyState Lib "User32" (ByVal nVirtKey As Long) As Integer
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'Lista de cabezas
Public Type tIndiceCabeza
    Head(1 To 4) As Integer
End Type

Public Type tIndiceCuerpo
    Body(1 To 4) As Integer
    HeadOffsetX As Integer
    HeadOffsetY As Integer
End Type

Public Type tIndiceFx
    Animacion As Integer
    OffsetX As Integer
    OffsetY As Integer
End Type



    Public Declare Function FindWindow Lib "User32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    
    Public Const WM_SETTEXT = &HC
    Public Const WM_GETTEXT = &HD
    Public Const WM_GETTEXTLENGTH = &HE
    Public Const EM_SETREADONLY = &HCF
    

Public Const DM_BITSPERPEL = &H40000
Public Const DM_PELSWIDTH = &H80000
Public Const DM_PELSHEIGHT = &H100000
Public Const CDS_UPDATEREGISTRY = &H1
Public Const CDS_TEST = &H4
Public Const DISP_CHANGE_SUCCESSFUL = 0
Public Const DISP_CHANGE_RESTART = 1
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long) As Long

