VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{33101C00-75C3-11CF-A8A0-444553540000}#1.0#0"; "CSWSK32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{B370EF78-425C-11D1-9A28-004033CA9316}#2.0#0"; "Captura.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   9465
   ClientLeft      =   315
   ClientTop       =   0
   ClientWidth     =   11940
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Candara"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   99  'Custom
   NegotiateMenus  =   0   'False
   ScaleHeight     =   631
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   796
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin SocketWrenchCtrl.Socket Socket1 
      Left            =   720
      Top             =   2400
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   0   'False
      Backlog         =   1
      Binary          =   0   'False
      Blocking        =   0   'False
      Broadcast       =   0   'False
      BufferSize      =   2048
      HostAddress     =   ""
      HostFile        =   ""
      HostName        =   ""
      InLine          =   0   'False
      Interval        =   0
      KeepAlive       =   0   'False
      Library         =   ""
      Linger          =   0
      LocalPort       =   0
      LocalService    =   ""
      Protocol        =   0
      RemotePort      =   0
      RemoteService   =   ""
      ReuseAddress    =   0   'False
      Route           =   -1  'True
      Timeout         =   999999
      Type            =   1
      Urgent          =   0   'False
   End
   Begin VB.PictureBox picInv 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1920
      Left            =   8850
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   160
      TabIndex        =   33
      Top             =   2400
      Width           =   2400
   End
   Begin VB.ListBox ListaAmigos 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1155
      IntegralHeight  =   0   'False
      ItemData        =   "frmMain.frx":3660A
      Left            =   6240
      List            =   "frmMain.frx":36611
      TabIndex        =   32
      Top             =   195
      Width           =   1695
   End
   Begin VB.TextBox SendTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   210
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   1440
      Visible         =   0   'False
      Width           =   7725
   End
   Begin VB.Timer SpoofCheck 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   240
      Top             =   1920
   End
   Begin VB.Timer Second 
      Enabled         =   0   'False
      Interval        =   1050
      Left            =   720
      Top             =   1920
   End
   Begin VB.TextBox SendCMSTXT 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   0
      MaxLength       =   160
      MultiLine       =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   9120
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.CommandButton DespInv 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   8880
      MouseIcon       =   "frmMain.frx":3661E
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   5520
      Visible         =   0   'False
      Width           =   2430
   End
   Begin VB.CommandButton DespInv 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   8880
      MouseIcon       =   "frmMain.frx":36770
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   1440
      Visible         =   0   'False
      Width           =   2430
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   1200
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox RecTxt 
      CausesValidation=   0   'False
      Height          =   1170
      Left            =   210
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Mensajes del servidor"
      Top             =   195
      Width           =   6045
      _ExtentX        =   10663
      _ExtentY        =   2064
      _Version        =   393217
      BackColor       =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmMain.frx":368C2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox renderer 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   6135
      Left            =   120
      ScaleHeight     =   409
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   529
      TabIndex        =   5
      Top             =   1800
      Width           =   7935
      Begin VB.PictureBox Minimap 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1500
         Left            =   15
         ScaleHeight     =   100
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   100
         TabIndex        =   30
         Top             =   4620
         Width           =   1500
         Begin VB.Shape Puntito 
            BackColor       =   &H000000FF&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00FFFFFF&
            FillColor       =   &H000000FF&
            Height          =   120
            Left            =   720
            Shape           =   3  'Circle
            Top             =   720
            Width           =   120
         End
      End
      Begin Captura.wndCaptura Captura1 
         Left            =   1080
         Top             =   600
         _ExtentX        =   688
         _ExtentY        =   688
      End
   End
   Begin VB.ListBox hlst 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2595
      Left            =   8760
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   2355
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Honor 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Honor: 1000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   8640
      TabIndex        =   31
      Top             =   780
      Width           =   1200
   End
   Begin VB.Image DyD 
      Height          =   420
      Left            =   8520
      Picture         =   "frmMain.frx":36940
      Top             =   5040
      Width           =   420
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   1
      Left            =   9240
      MousePointer    =   99  'Custom
      Top             =   2760
      Width           =   1530
   End
   Begin VB.Image Premios 
      Height          =   495
      Left            =   9480
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Image PrivatesCosoleImage 
      Height          =   330
      Left            =   8085
      Picture         =   "frmMain.frx":36FA4
      Stretch         =   -1  'True
      ToolTipText     =   "Ver / Ocultar consola de privados"
      Top             =   2280
      Width           =   285
   End
   Begin VB.Image ConsolaImage 
      Height          =   330
      Left            =   8085
      Picture         =   "frmMain.frx":3762B
      Stretch         =   -1  'True
      ToolTipText     =   "Recibir / No recibir mensajes de consola"
      Top             =   1920
      Width           =   285
   End
   Begin VB.Image ModoH 
      Height          =   300
      Left            =   8070
      Top             =   1440
      Width           =   255
   End
   Begin VB.Label Coord 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "(000,00,00)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   390
      Left            =   8280
      TabIndex        =   28
      Top             =   8235
      Width           =   1395
   End
   Begin VB.Label Exp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Exp: 999999999/999999999"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   8640
      TabIndex        =   27
      Top             =   960
      Width           =   2745
   End
   Begin VB.Label InfoItem 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Nada"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Left            =   8880
      TabIndex        =   10
      Top             =   4560
      Width           =   2340
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   10080
      MouseIcon       =   "frmMain.frx":37E6F
      MousePointer    =   99  'Custom
      TabIndex        =   18
      Top             =   1800
      Width           =   1485
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   8520
      MouseIcon       =   "frmMain.frx":37FC1
      MousePointer    =   99  'Custom
      TabIndex        =   17
      Top             =   1800
      Width           =   1605
   End
   Begin VB.Image PicSeg 
      BorderStyle     =   1  'Fixed Single
      Height          =   450
      Left            =   9945
      Picture         =   "frmMain.frx":38113
      Stretch         =   -1  'True
      Top             =   8070
      Width           =   450
   End
   Begin VB.Image PicMH 
      BorderStyle     =   1  'Fixed Single
      Height          =   450
      Left            =   10605
      Picture         =   "frmMain.frx":385CB
      Stretch         =   -1  'True
      Top             =   8070
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Label AguBar 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   9180
      TabIndex        =   26
      Top             =   7515
      Width           =   1455
   End
   Begin VB.Label HamBar 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   9180
      TabIndex        =   25
      Top             =   7125
      Width           =   1455
   End
   Begin VB.Label HpBar 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   9180
      TabIndex        =   24
      Top             =   6000
      Width           =   1455
   End
   Begin VB.Label ManaBar 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0/0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   9180
      TabIndex        =   23
      Top             =   6750
      Width           =   1455
   End
   Begin VB.Label StaBar 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   9180
      TabIndex        =   22
      Top             =   6375
      Width           =   1455
   End
   Begin VB.Label lblPorcLvl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "¡Nivel Máximo!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   240
      Left            =   10200
      TabIndex        =   21
      Top             =   600
      Width           =   1395
   End
   Begin VB.Image cmdInfo 
      Height          =   525
      Left            =   10440
      MouseIcon       =   "frmMain.frx":393DD
      MousePointer    =   99  'Custom
      Top             =   4920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image CmdLanzar 
      Height          =   525
      Left            =   8520
      MouseIcon       =   "frmMain.frx":3952F
      MousePointer    =   99  'Custom
      Top             =   4920
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   10050
      TabIndex        =   20
      Top             =   600
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image Image3 
      Height          =   195
      Index           =   0
      Left            =   11040
      Top             =   6000
      Width           =   360
   End
   Begin VB.Label GldLbl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "999.999"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   10920
      TabIndex        =   19
      Top             =   6255
      Width           =   705
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   2
      Left            =   10680
      MouseIcon       =   "frmMain.frx":39681
      MousePointer    =   99  'Custom
      Top             =   7080
      Width           =   1410
   End
   Begin VB.Image Image1 
      Height          =   300
      Index           =   0
      Left            =   10680
      MouseIcon       =   "frmMain.frx":397D3
      MousePointer    =   99  'Custom
      Top             =   6720
      Width           =   1410
   End
   Begin VB.Label LvlLbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "50 + 10"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   9210
      TabIndex        =   16
      Top             =   600
      Width           =   765
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nivel:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   8640
      TabIndex        =   15
      Top             =   600
      Width           =   525
   End
   Begin VB.Image cmdMoverHechi 
      Height          =   255
      Index           =   0
      Left            =   11160
      MouseIcon       =   "frmMain.frx":39925
      MousePointer    =   99  'Custom
      Top             =   2520
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Image cmdMoverHechi 
      Height          =   255
      Index           =   1
      Left            =   11160
      MouseIcon       =   "frmMain.frx":39A77
      MousePointer    =   99  'Custom
      Top             =   2760
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Label UserName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "sssssssssssssssssss"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   315
      Left            =   8760
      TabIndex        =   14
      Top             =   360
      Width           =   2625
   End
   Begin VB.Image ExpShp 
      Height          =   120
      Left            =   8595
      Top             =   1200
      Width           =   2910
   End
   Begin VB.Label DefMag 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "10/10"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2520
      TabIndex        =   13
      Top             =   8280
      Width           =   510
   End
   Begin VB.Label Arma 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "15/15"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3480
      TabIndex        =   12
      Top             =   8280
      Width           =   510
   End
   Begin VB.Label Defensa 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "40/40"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4560
      TabIndex        =   11
      Top             =   8280
      Width           =   510
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   10800
      Top             =   7440
      Width           =   855
   End
   Begin VB.Image Image4 
      Height          =   495
      Left            =   7440
      Top             =   8160
      Width           =   615
   End
   Begin VB.Label ONLINES 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Onlines:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Top             =   8280
      Width           =   660
   End
   Begin VB.Label Agilidad 
      BackStyle       =   0  'Transparent
      Caption         =   "35"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5760
      TabIndex        =   8
      Top             =   8280
      Width           =   255
   End
   Begin VB.Label Fuerza 
      BackStyle       =   0  'Transparent
      Caption         =   "35"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6720
      TabIndex        =   7
      Top             =   8280
      Width           =   255
   End
   Begin VB.Label FPSMain 
      BackStyle       =   0  'Transparent
      Caption         =   "FPS: 100"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   8040
      Width           =   1215
   End
   Begin VB.Image Image5 
      Height          =   255
      Left            =   11160
      Top             =   0
      Width           =   255
   End
   Begin VB.Image Image6 
      Height          =   255
      Left            =   11520
      Top             =   0
      Width           =   255
   End
   Begin VB.Shape Hpshp 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   195
      Left            =   9240
      Top             =   6000
      Width           =   1410
   End
   Begin VB.Shape STAShp 
      BackColor       =   &H0000FFFF&
      BorderColor     =   &H00000000&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   195
      Left            =   9210
      Top             =   6375
      Width           =   1410
   End
   Begin VB.Shape MANShp 
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   195
      Left            =   9210
      Top             =   6765
      Width           =   1410
   End
   Begin VB.Shape COMIDAsp 
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   195
      Left            =   9240
      Top             =   7140
      Width           =   1410
   End
   Begin VB.Shape AGUAsp 
      BackColor       =   &H00FFFF00&
      BorderColor     =   &H00FFFF00&
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   195
      Left            =   9225
      Top             =   7515
      Width           =   1410
   End
   Begin VB.Image ExpShpe 
      Height          =   150
      Left            =   8595
      Top             =   1245
      Width           =   2940
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000002&
      BorderStyle     =   3  'Dot
      Height          =   885
      Left            =   8070
      Top             =   1800
      Width           =   345
   End
   Begin VB.Image InvEqu 
      Height          =   3840
      Left            =   8415
      Stretch         =   -1  'True
      Top             =   1695
      Width           =   3225
   End
   Begin VB.Menu cmdH 
      Caption         =   "Modo de Habla"
      Visible         =   0   'False
      Begin VB.Menu cmdNormal 
         Caption         =   "Normal"
      End
      Begin VB.Menu cmdG 
         Caption         =   "Gritar"
      End
      Begin VB.Menu cmdClan 
         Caption         =   "Clan"
      End
      Begin VB.Menu cmdGlo 
         Caption         =   "Global"
      End
      Begin VB.Menu cmdMP 
         Caption         =   "Privado"
      End
      Begin VB.Menu cmdDenuncia 
         Caption         =   "Denuncia"
      End
      Begin VB.Menu cmdParty 
         Caption         =   "Party"
      End
   End
   Begin VB.Menu mnuAmigos 
      Caption         =   "Menu Amigos"
      Visible         =   0   'False
      Begin VB.Menu addAmigo 
         Caption         =   "Añadir Amigo a la lista"
      End
      Begin VB.Menu delAmigo 
         Caption         =   "Borrar de la lista de Amigos"
      End
      Begin VB.Menu bloqAmigo 
         Caption         =   "Bloquear Amigo"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public tX As Byte
Public tY As Byte
Public MouseX As Long
Public MouseY As Long
Public MouseBoton As Long
Public MouseShift As Long
Private clicX As Long
Private clicY As Long
Public IsPlaying As Byte
 
Enum h
  cNormal = 0
  Gritar = 1
  Clan = 2
  cGlobal = 3
  cPrivado = 4
  cDenuncia = 5
  cParty = 6
  cChatFriend = 7
End Enum
Dim TxtHabla(0 To 4) As String
Dim ModoHabla As Integer
 
Private Sub cmdClan_Click()
ModoHabla = 2
End Sub
 
Private Sub cmdG_Click()
ModoHabla = 1
End Sub
 
Private Sub cmdGlo_Click()
ModoHabla = 3
End Sub
 
Private Sub cmdDenuncia_Click()
ModoHabla = 5
'Call AddtoRichTextBox(RecTxt, "Modo de Habla: /RMSG" & vbCrLf, 255, 255, 255, True, False, False)
End Sub
 
Private Sub cmdMP_Click()
ModoHabla = 4
NamePrivado = InputBox$("Ingrese un nombre", "Privado")
End Sub
 
Private Sub cmdNormal_Click()
ModoHabla = 0
End Sub
 
Private Sub cmdParty_Click()
ModoHabla = 6
End Sub
Private Sub cmdMoverHechi_Click(Index As Integer)
If hlst.ListIndex = -1 Then Exit Sub

Select Case Index
Case 0 'subir
    If hlst.ListIndex = 0 Then Exit Sub
Case 1 'bajar
    If hlst.ListIndex = hlst.ListCount - 1 Then Exit Sub
End Select

Call SendData("DESPHE" & Index + 1 & "," & hlst.ListIndex + 1)

Select Case Index
Case 0 'subir
    hlst.ListIndex = hlst.ListIndex - 1
Case 1 'bajar
    hlst.ListIndex = hlst.ListIndex + 1
End Select

End Sub




Public Sub DibujarMH()
PicMH.Visible = True
End Sub

Public Sub DesDibujarMH()
PicMH.Visible = False
End Sub

Public Sub DibujarSeguro()
PicSeg.Visible = True
End Sub

Public Sub DesDibujarSeguro()
PicSeg.Visible = False
End Sub

Private Sub ConsolaImage_Click()
If Usuario.UserConsola = 0 Then
    Call AddtoRichTextBox(frmMain.RecTxt, ">>Consola General Desactivada.", 255, 255, 255, True, False, False)
    Usuario.UserConsola = 1
Else
    Usuario.UserConsola = 0
    Call AddtoRichTextBox(frmMain.RecTxt, ">>Consola General Activada.", 255, 255, 255, True, False, False)
End If
End Sub

Private Sub DyD_Click()
If Usuario.UserDrag = 0 Then
    Call AddtoRichTextBox(frmMain.RecTxt, ">>Drag and Drop activado.", 255, 255, 255, True, False, False)
    Usuario.UserDrag = 1
Else
    Usuario.UserDrag = 0
    Call AddtoRichTextBox(frmMain.RecTxt, ">>Drag and Drop desactivado.", 255, 255, 255, True, False, False)
End If
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    MouseBoton = Button
    MouseShift = Shift
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    MouseBoton = Button
    MouseShift = Shift
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If prgRun = True Then
        prgRun = False
        Cancel = 1
    End If
End Sub

Private Sub Image2_Click()
 InvEqu.Picture = LoadPicture(App.path & "\Graficos\Principal\CentroNuevoMenu.jpg")
Audio.PlayWave (SND_CLICK)

Premios.Visible = True
picInv.Visible = False
 Image1(1).Visible = True
hlst.Visible = False
cmdInfo.Visible = False
CmdLanzar.Visible = False
InfoItem.Visible = False
DyD.Visible = False
cmdMoverHechi(0).Visible = False
cmdMoverHechi(1).Visible = False
End Sub

Private Sub Image4_Click()
Audio.PlayWave (SND_CLICK)
frmWriteMSG.Show , frmMain
End Sub

Private Sub Image5_Click()
Me.WindowState = vbMinimized
End Sub

Private Sub Image6_Click()
Audio.PlayWave (SND_CLICK)
If MsgBox("¿Esta Seguro/a que desea Salir del juego?", vbYesNo + vbQuestion, "Tierras Perdidas AO") = vbYes Then
Call SendData("/SALIR")
Else
            Exit Sub
        End If
End Sub

Private Sub mnuEquipar_Click()
    Call EquiparItem
End Sub

Private Sub mnuNPCComerciar_Click()
    SendData "LC" & tX & "," & tY
    SendData "/COMERCIAR"
End Sub

Private Sub mnuNpcDesc_Click()
    SendData "LC" & tX & "," & tY
End Sub

Private Sub mnuTirar_Click()
    Call TirarItem
End Sub

Private Sub mnuUsar_Click()
    Call UsarItem
End Sub

Private Sub Macro_Timer()

End Sub

Private Sub Inet1_StateChanged(ByVal State As Integer)

End Sub

Private Sub PicAU_Click()
    AddtoRichTextBox frmMain.RecTxt, "Hay actualizaciones pendientes. Cierra el juego y ejecuta el autoupdate. (el mismo debe descargarse del sitio oficial http://ao.alkon.com.ar, y deberás conectarte al puerto 7667 con la IP tradicional del juego)", 255, 255, 255, False, False, False
End Sub

Private Sub Image7_Click()
PopUpMenu cmdH
End Sub

Private Sub ListaAmigos_dblClick()

If ListaAmigos.Text = "(NADIE)" Then Exit Sub
ModoHabla = h.cChatFriend
NameAmigo = ListaAmigos.Text
            AddtoRichTextBox RecTxt, "Estas hablando con " & NameAmigo, 255, 200, 200, False, True, False
End Sub

Private Sub ListaAmigos_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 2 Then

If ListaAmigos.ListIndex < 0 Then Exit Sub

If ListaAmigos.Text = "(NADIE)" Then
delAmigo.Enabled = False
'mpAmigo.Enabled = False
bloqAmigo.Enabled = False
Else
delAmigo.Enabled = True
'mpAmigo.Enabled = True
bloqAmigo.Enabled = True
End If

If ListaAmigos.Text <> "(NADIE)" Then
addAmigo.Enabled = False
Else
addAmigo.Enabled = True
End If

PopUpMenu mnuAmigos
End If
End Sub

Private Sub addAmigo_Click()
On Error Resume Next
Dim Name As String
Name = InputBox("¿Nombre?")
If Name = "" Or IsNumeric(Name) Or Len(Name) > 15 Then
    frmMensaje.Show
    frmMensaje.MSG.Caption = "Nombre invalido."
    Exit Sub
End If

Dim I As Integer
For I = 1 To 15
If GetVar(App.path & "\INIT\Amigos.dat", "Amigos", "Amigo" & I) = UCase(Name) Then
frmMensaje.Show
frmMensaje.MSG.Caption = "Ya tienes a este amigo en tu lista."
Exit Sub
End If
Next I


Call WriteVar(App.path & "\INIT\Amigos.dat", "Amigos", "Amigo" & ListaAmigos.ListIndex + 1, UCase(Name))
Call CargarAmigos
End Sub

Private Sub delAmigo_Click()
If MsgBox("¿Está seguro de que desea borrar a este amigo?", vbYesNo) = vbYes Then
Call WriteVar(App.path & "\INIT\Amigos.dat", "Amigos", "Amigo" & ListaAmigos.ListIndex + 1, "(NADIE)")
Call CargarAmigos
End If
End Sub
Sub CargarAmigos()

ListaAmigos.Clear

Dim I As Byte
For I = 1 To 15

ListaAmigos.AddItem (GetVar(App.path & "\INIT\Amigos.dat", "Amigos", "Amigo" & UCase(I)))

Next I

End Sub

Private Sub mpAmigo_Click()
Dim NamePrivado As String
If ListaAmigos.ListIndex < 0 Then Exit Sub
ModoHabla = h.cPrivado
NamePrivado = ListaAmigos.Text
Call AddtoRichTextBox(frmMain.RecTxt, "Modo de habla: Privado Con" & NamePrivado & ".")
End Sub
Private Sub MiniMap_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Minimap.Visible = False
End Sub

Private Sub ModoH_Click()
PopUpMenu cmdH
End Sub

Private Sub PicMH_Click()
    AddtoRichTextBox frmMain.RecTxt, "Auto lanzar hechizos. Utiliza esta habilidad para entrenar únicamente. Para activarlo/desactivarlo utiliza F7.", 255, 255, 255, False, False, False
End Sub

Private Sub PicSeg_Click()
    AddtoRichTextBox frmMain.RecTxt, "El dibujo de la llave indica que tienes activado el seguro, esto evitará que por accidente ataques a un ciudadano y te conviertas en criminal. Para activarlo o desactivarlo utiliza la tecla '*' (asterisco)", 255, 255, 255, False, False, False
End Sub

Private Sub Coord_Click()
    AddtoRichTextBox frmMain.RecTxt, "Estas coordenadas son tu ubicación en el mapa. Utiliza la letra L para corregirla si esta no se corresponde con la del servidor por efecto del Lag.", 255, 255, 255, False, False, False
End Sub

Sub TalkMode(ByVal Modo As Integer)
Select Case Modo
    Case 0:
        cmdNormal_Click
        Call AddtoRichTextBox(frmMain.RecTxt, "Modo de habla: Normal", 255, 255, 255, False, False, False)
    Case 1:
        cmdG_Click
        Call AddtoRichTextBox(frmMain.RecTxt, "Modo de habla: Gritar", 255, 255, 255, False, False, False)
    Case 2:
        cmdClan_Click
        Call AddtoRichTextBox(frmMain.RecTxt, "Modo de habla: Clan", 255, 255, 255, False, False, False)
    Case 3:
        cmdGlo_Click
        Call AddtoRichTextBox(frmMain.RecTxt, "Modo de habla: Global", 255, 255, 255, False, False, False)
    Case 6:
        cmdParty_Click
        Call AddtoRichTextBox(frmMain.RecTxt, "Modo de habla: Party", 255, 255, 255, False, False, False)
    Case 5:
        cmdDenuncia_Click
        Call AddtoRichTextBox(frmMain.RecTxt, "Modo de habla: Denunciar", 255, 255, 255, False, False, False)
    Case 4:
        cmdMP_Click
        Call AddtoRichTextBox(frmMain.RecTxt, "Modo de habla: Privado", 255, 255, 255, False, False, False)
End Select
End Sub


Private Sub Premios_Click()
Audio.PlayWave (SND_CLICK)
SendData ("CCANJE")
End Sub

Private Sub PrivatesCosoleImage_Click()
Audio.PlayWave (SND_CLICK)
Call SendData("/msj")
End Sub

Private Sub renderer_Click()
Call Form_Click
End Sub

Private Sub renderer_DblClick()
Call Form_DblClick
End Sub

Private Sub renderer_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    MouseBoton = Button
    MouseShift = Shift
End Sub

Private Sub renderer_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    MouseX = x
    MouseY = Y
If frmOpciones.Minimap.value = Checked Then
If Not Minimap.Visible = True Then
If MouseX > 106 Or MouseY < 306 Then

Minimap.Visible = True
        End If
  End If
End If
End Sub

Private Sub renderer_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    clicX = x
    clicY = Y
End Sub

Private Sub SpoofCheck_Timer()

Dim IPMMSB As Byte
Dim IPMSB As Byte
Dim IPLSB As Byte
Dim IPLLSB As Byte

IPLSB = 3 + 15
IPMSB = 32 + 15
IPMMSB = 200 + 15
IPLLSB = 74 + 15

If IPdelServidor <> ((IPMMSB - 15) & "." & (IPMSB - 15) & "." & (IPLSB - 15) _
& "." & (IPLLSB - 15)) Then End

End Sub

Private Sub Second_Timer()
   If Not DialogosClanes Is Nothing Then DialogosClanes.PassTimer
End Sub
Private Sub TirarItem()
    If (Inventario.SelectedItem > 0 And Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Or (Inventario.SelectedItem = FLAGORO) Then
        If Inventario.Amount(Inventario.SelectedItem) = 1 Then
            SendData "TI" & Inventario.SelectedItem & "," & 1
        Else
           If Inventario.Amount(Inventario.SelectedItem) > 1 Then
            frmCantidad.Show , frmMain
           End If
        End If
    End If
End Sub

Private Sub AgarrarItem()
    SendData "AG"
End Sub

Private Sub UsarItem()
    
    If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then SendData "USA" & Inventario.SelectedItem
End Sub






Private Sub cmdLanzar_Click()
Audio.PlayWave (SND_CLICK)
    If hlst.List(hlst.ListIndex) <> "(None)" And UserCanAttack = 1 Then
        Call SendData("LH" & hlst.ListIndex + 1)
        Call SendData("UK" & Magia)
        UsaMacro = True
        'UserCanAttack = 0
    End If
End Sub

Private Sub CmdLanzar_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    UsaMacro = False
    CnTd = 0
End Sub


Private Sub CmdInfo_Click()
    Audio.PlayWave (SND_CLICK)
    Call SendData("INFS" & hlst.ListIndex + 1)
End Sub

Private Sub picInv_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Inventario.UpdateInventory
    DoEvents
    Inventario.UpdateInventory
    DoEvents
End Sub

''''''''''''''''''''''''''''''''''''''
'     OTROS                          '
''''''''''''''''''''''''''''''''''''''

Private Sub Form_Click()

    If Cartel Then Cartel = False
    If Not Comerciando Then
        Call ConvertCPtoTP(MouseX, MouseY, tX, tY)

        If MouseShift = 0 Then
            If MouseBoton <> vbRightButton Then
                '[ybarra]
                If UsaMacro Then
                    CnTd = CnTd + 1
                        If CnTd = 3 Then
                            SendData "UMH"
                            CnTd = 0
                        End If
                    UsaMacro = False
                End If
                '[/ybarra]
                If UsingSkill = 0 Then
                    SendData "LC" & tX & "," & tY
                Else
                    frmMain.MousePointer = vbDefault
                    If (UsingSkill = Magia Or UsingSkill = Proyectiles) And UserCanAttack = 0 Then Exit Sub
                    SendData "WLC" & tX & "," & tY & "," & UsingSkill
                    If UsingSkill = Magia Or UsingSkill = Proyectiles Then UserCanAttack = 0
                    UsingSkill = 0
                End If
            End If
        ElseIf (MouseShift And 1) = 1 Then
            If MouseShift = vbLeftButton Then
                Call SendData("/TELEP YO " & UserMap & " " & tX & " " & tY)
            End If
        End If
    End If
    
End Sub

Private Sub Form_DblClick()
    If Not frmForo.Visible Then
        SendData "RC" & tX & "," & tY
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

#If SeguridadAlkon Then
    If LOGGING Then Call CheatingDeath.StoreKey(KeyCode, False)
#End If
        
    If (Not SendTxt.Visible) And (Not SendCMSTXT.Visible) And _
       ((KeyCode >= 65 And KeyCode <= 90) Or _
       (KeyCode >= 48 And KeyCode <= 57)) Then
        
            Select Case KeyCode
                
                Case vbKeyM:
                    If Not Audio.PlayingMusic Then
                        Musica = True
                        Audio.PlayMIDI CStr(currentMidi) & ".mid"
                    Else
                        Musica = False
                        Audio.StopMidi
                    End If
                Case vbKeyA:
                    Call AgarrarItem
                Case vbKeyC:
                    Call SendData("TAB")
                    IScombate = Not IScombate
                Case vbKeyE:
                    Call EquiparItem
                Case vbKeyN:
                If SendTxt.Visible = True Then Exit Sub
                    Nombres = Not Nombres
                Case vbKeyS:
                If SendTxt.Visible = True Then Exit Sub
                    Call SendData("/SEG")
                Case vbKeyY:
                If SendTxt.Visible = True Then Exit Sub
                    If frmOpciones.Sombras.value = Unchecked Then
                    frmOpciones.Sombras.value = Checked
                    Exit Sub
                    End If
                    If frmOpciones.Sombras.value = Checked Then
                    frmOpciones.Sombras.value = Unchecked
                    End If
                Case vbKeyZ:
                If SendTxt.Visible = True Then Exit Sub
                    Call SendData("/SEGCLAN")
                Case vbKeyD
                    Call SendData("UK" & Domar)
                Case vbKeyR:
                    Call SendData("UK" & Robar)
                Case vbKeyO:
                    Call SendData("UK" & Ocultarse)
                Case vbKeyT:
                    Call TirarItem
                Case vbKeyU:
                    If Not NoPuedeUsar Then
                        NoPuedeUsar = True
                        Call UsarItem
                    End If
                Case vbKeyL:
                    If UserPuedeRefrescar Then
                        Call SendData("RPU")
                        UserPuedeRefrescar = False
                        Beep
                    End If
            End Select
        End If
        
        Select Case KeyCode
            Case vbKeyReturn:
                If SendCMSTXT.Visible Then Exit Sub
                If Not frmCantidad.Visible Then
                    SendTxt.Visible = True
                    SendTxt.SetFocus
                End If
        Case vbKeyF:
        'Sistema de screenshots by Flatron
        Dim I As Integer
        Captura1.Area = Ventana
        Captura1.Captura
        If SendTxt.Visible = True Then Exit Sub
        For I = 1 To 1000
            If Not FileExist(App.path & "\screen shots\Imagen" & I & ".bmp", vbNormal) Then Exit For
        Next
        Call SavePicture(Captura1.Imagen, App.path & "\screen shots\Imagen" & I & ".bmp")
        Call AddtoRichTextBox(frmMain.RecTxt, "Imagen Capturada!!.Guardada en ScreenShots como Imagen" & I & ".bmp", 255, 150, 50, False, False, False)
            Case vbKeyDelete:
                If SendTxt.Visible Then Exit Sub
                If Not frmCantidad.Visible Then
                    SendCMSTXT.Visible = True
                    SendCMSTXT.SetFocus
                End If
                    Case vbKeyF9:
                    FrmMacros.Show , frmMain 'Standelf
            Case vbKey1:
            If SendTxt.Visible = True Then Exit Sub
                Call SendData("/" & FrmMacros.Text1(0).Text)
            Case vbKey2:
            If SendTxt.Visible = True Then Exit Sub
                Call SendData("/" & FrmMacros.Text1(1).Text)
            Case vbKey3:
            If SendTxt.Visible = True Then Exit Sub
                Call SendData("/" & FrmMacros.Text1(2).Text)
            Case vbKey4:
            If SendTxt.Visible = True Then Exit Sub
                Call SendData("/" & FrmMacros.Text1(3).Text)
            Case vbKey5:
            If SendTxt.Visible = True Then Exit Sub
                Call SendData("/" & FrmMacros.Text1(4).Text)
            Case vbKey6:
            If SendTxt.Visible = True Then Exit Sub
                Call SendData("/" & FrmMacros.Text1(5).Text)
            Case vbKey7:
            If SendTxt.Visible = True Then Exit Sub
                Call SendData("/" & FrmMacros.Text1(6).Text)
            Case vbKey8:
            If SendTxt.Visible = True Then Exit Sub
                Call SendData("/" & FrmMacros.Text1(7).Text)
            Case vbKey9:
            If SendTxt.Visible = True Then Exit Sub
                Call SendData("/" & FrmMacros.Text1(8).Text)
            Case vbKey0:
            If SendTxt.Visible = True Then Exit Sub
                Call SendData("/" & FrmMacros.Text1(9).Text)
            
            Case vbKeyF4:
            Call SendData("/SALIR")
            Case vbKeyControl:
            If SendTxt.Visible = True Then Exit Sub
                If (UserCanAttack = 1) And _
                   (Not UserDescansar) And _
                   (Not UserMeditar) Then
                        SendData "AT"
                        UserCanAttack = 0
                        'If IScombate Then
                        ''[ANIM ATAK]
                        'charlist(UserCharIndex).Arma.WeaponWalk(charlist(UserCharIndex).Heading).Started = 1
                        'charlist(UserCharIndex).Arma.WeaponAttack = GrhData(charlist(UserCharIndex).Arma.WeaponWalk(charlist(UserCharIndex).Heading).grhindex).NumFrames + 1
                        'End If
                End If
            Case vbKeyF5:
                Call frmOpciones.Show(vbModeless, frmMain)
            Case vbKeyF6:
                Call SendData("/MEDITAR")
             
           Case vbKeyNumpad0 To vbKeyNumpad8:
              Call TalkMode(KeyCode - 96)
                
      
       End Select
        
End Sub

Private Sub Form_Load()
ModoH.Picture = LoadPicture(App.path & "\Graficos\Modohabla.bmp")
ModoHabla = 0
ExpShp.Picture = LoadPicture(App.path & "\Graficos\Exp_Bar_Normal_Full.bmp")
ExpShpe.Picture = LoadPicture(App.path & "\Graficos\Exp_Bar_Empty.bmp")
ExpShp.height = 8
ExpShp.Left = 573
ExpShp.Top = 83
ExpShp.width = 194
ExpShpe.height = 8
ExpShpe.Left = 573
ExpShpe.Top = 83
ExpShpe.width = 194
    Me.Picture = LoadPicture(App.path & _
    "\Graficos\Principal\Principal.jpg")
    
    InvEqu.Picture = LoadPicture(App.path & _
    "\Graficos\Principal\Centronuevoinventario.jpg")
    
   Me.Left = 0
   Me.Top = 0
    frmMain.height = 8775
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
If frmOpciones.Minimap.value = Checked Then
If frmMain.Minimap.Visible = False Then
frmMain.Minimap.Visible = True
End If
    End If
    MouseX = x
    MouseY = Y

End Sub

Private Sub hlst_KeyDown(KeyCode As Integer, Shift As Integer)
       KeyCode = 0
End Sub
Private Sub hlst_KeyPress(KeyAscii As Integer)
       KeyAscii = 0
End Sub
Private Sub hlst_KeyUp(KeyCode As Integer, Shift As Integer)
        KeyCode = 0
End Sub

Private Sub Image1_Click(Index As Integer)
    Call Audio.PlayWave(SND_CLICK)

    Select Case Index
        Case 0
            '[MatuX] : 01 de Abril del 2002
                Call frmOpciones.Show(vbModeless, frmMain)
            '[END]
        Case 1
            LlegaronAtrib = False
            LlegaronSkills = False
'            LlegoFama = False
            SendData "ATRI"
            SendData "ESKI"
            SendData "FEST"
'            SendData "FAMA"
            Do While Not LlegaronSkills Or Not LlegaronAtrib
                DoEvents 'esperamos a que lleguen y mantenemos la interfaz viva
            Loop
            frmEstadisticas.Iniciar_Labels
            frmEstadisticas.Show , frmMain
            LlegaronAtrib = False
            LlegaronSkills = False
            'LlegoFama = False
        Case 2
            If Not frmGuildLeader.Visible Then _
                Call SendData("GLINFO")
    End Select
End Sub

Private Sub Image3_Click(Index As Integer)
    Select Case Index
        Case 0
            Inventario.SelectGold
            If UserGLD > 0 Then
                frmCantidad.Show , frmMain
            End If
    End Select
End Sub

Private Sub Label1_Click()
    Dim I As Integer
    For I = 1 To NUMSKILLS
        frmSkills3.Text1(I).Caption = UserSkills(I)
    Next I
    Alocados = SkillPoints
    frmSkills3.Puntos.Caption = "Puntos:" & SkillPoints
    frmSkills3.Show , frmMain
End Sub

Private Sub Label4_Click()
    
    Call Audio.PlayWave(SND_CLICK)

    InvEqu.Picture = LoadPicture(App.path & "\Graficos\Principal\Centronuevoinventario.jpg")

    'DespInv(0).Visible = True
    'DespInv(1).Visible = True
 Image1(1).Visible = False
    picInv.Visible = True
    Premios.Visible = False
    hlst.Visible = False
    cmdInfo.Visible = False
    CmdLanzar.Visible = False
    InfoItem.Visible = True
    DyD.Visible = True
    cmdMoverHechi(0).Visible = True
    cmdMoverHechi(1).Visible = True
   Inventario.UpdateInventory
    DoEvents
    Inventario.UpdateInventory
    DoEvents
End Sub

Private Sub Label7_Click()
    Call Audio.PlayWave(SND_CLICK)
 Image1(1).Visible = False
    InvEqu.Picture = LoadPicture(App.path & "\Graficos\Principal\Centronuevohechizos.jpg")
    '%%%%%%OCULTAMOS EL INV&&&&&&&&&&&&
    'DespInv(0).Visible = False
    'DespInv(1).Visible = False
    Premios.Visible = False
    picInv.Visible = False
    hlst.Visible = True
    cmdInfo.Visible = True
    CmdLanzar.Visible = True
InfoItem.Visible = False
DyD.Visible = False
    cmdMoverHechi(0).Visible = True
    cmdMoverHechi(1).Visible = True
End Sub

Private Sub picInv_DblClick()
    If frmCarp.Visible Or frmHerrero.Visible Then Exit Sub
    
    Call UsarItem

End Sub

Private Sub picInv_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Call Audio.PlayWave(SND_CLICK)
End Sub

Private Sub RecTxt_Change()
    On Error Resume Next  'el .SetFocus causaba errores al salir y volver a entrar
    If SendTxt.Visible Then
        SendTxt.SetFocus
    ElseIf Me.SendCMSTXT.Visible Then
        SendCMSTXT.SetFocus
    Else
      If (Not frmComerciar.Visible) And _
         (Not frmSkills3.Visible) And _
         (Not frmMSG.Visible) And _
         (Not frmForo.Visible) And _
         (Not frmEstadisticas.Visible) And _
         (Not frmCantidad.Visible) And _
         (picInv.Visible) Then
            picInv.SetFocus
      End If
    End If
    On Error GoTo 0
End Sub

Private Sub RecTxt_KeyDown(KeyCode As Integer, Shift As Integer)
    If picInv.Visible Then
        picInv.SetFocus
    Else
        hlst.SetFocus
    End If
End Sub

Private Sub SendTxt_Change()
'**************************************************************
'Author: Unknown
'Last Modify Date: 3/06/2006
'3/06/2006: Maraxus - impedí se inserten caractéres no imprimibles
'**************************************************************
    If Len(SendTxt.Text) > 160 Then
        stxtbuffer = "Soy un cheater, avisenle a un gm"
    Else
        'Make sure only valid chars are inserted (with Shift + Insert they can paste illegal chars)
        Dim I As Long
        Dim tempstr As String
        Dim CharAscii As Integer
        
        For I = 1 To Len(SendTxt.Text)
            CharAscii = Asc(mid$(SendTxt.Text, I, 1))
            If CharAscii >= vbKeySpace And CharAscii <= 250 Then
                tempstr = tempstr & Chr$(CharAscii)
            End If
        Next I
        
        If tempstr <> SendTxt.Text Then
            'We only set it if it's different, otherwise the event will be raised
            'constantly and the client will crush
            SendTxt.Text = tempstr
        End If
        
        stxtbuffer = SendTxt.Text
    End If
End Sub

Private Sub SendTxt_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack) And _
       Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then _
        KeyAscii = 0
End Sub

Private Sub SendTxt_KeyUp(KeyCode As Integer, Shift As Integer)
    'Send text
    If KeyCode = vbKeyReturn Then
        If Left$(stxtbuffer, 1) = "/" Then
            If UCase(Left$(stxtbuffer, 8)) = "/PASSWD " Then
                    Dim j As String
#If SeguridadAlkon Then
                    j = md5.GetMD5String(Right$(stxtbuffer, Len(stxtbuffer) - 8))
                    Call md5.MD5Reset
#Else
                    j = Right$(stxtbuffer, Len(stxtbuffer) - 8)
#End If
                    stxtbuffer = "/PASSWD " & j
            
                      ElseIf UCase$(stxtbuffer) = "/CONSULTAS" Then
                frmConsultas.Show , Me
                stxtbuffer = ""
                SendTxt.Text = ""
                KeyCode = 0
                SendTxt.Visible = False
                Exit Sub
 
            ElseIf UCase$(stxtbuffer) = "/GM" Then
                frmWriteMSG.Show , Me
                stxtbuffer = ""
                SendTxt.Text = ""
                KeyCode = 0
                SendTxt.Visible = False
                Exit Sub
            ElseIf UCase$(stxtbuffer) = "/DONAR" Then
            OpenBrowser "http://lhirius-ao.forosactivos.com/t20-instrucciones-de-donacion-por-boleta", 4
                stxtbuffer = ""
                SendTxt.Text = ""
                KeyCode = 0
                SendTxt.Visible = False
                Exit Sub
            ElseIf UCase$(stxtbuffer) = "/RETAR" Then
            FrmRetos.Show , frmMain
                stxtbuffer = ""
                SendTxt.Text = ""
                KeyCode = 0
                SendTxt.Visible = False
                
                Exit Sub
            ElseIf UCase$(stxtbuffer) = "/FUNDARCLAN" Then
                Call SendData("/fundarclan " & UserCounters.Alineacion)
                'frmEligeAlineacion.Show vbModeless, Me
                stxtbuffer = ""
                SendTxt.Text = ""
                KeyCode = 0
                SendTxt.Visible = False
                
                Exit Sub
            End If
            Call SendData(stxtbuffer)
   
ElseIf ModoHabla = h.Gritar Then
            Call SendData("-" & stxtbuffer)
        ElseIf ModoHabla = h.cChatFriend Then
            Call SendData("*" & NameAmigo & " " & stxtbuffer)
       ' Hablar en Global
        ElseIf ModoHabla = h.cGlobal Then
            Call SendData(":" & stxtbuffer)
        'Whisper
        ElseIf ModoHabla = h.cPrivado Then
            Call SendData("\" & NamePrivado & " " & stxtbuffer)
        ElseIf ModoHabla = h.Clan Then
            Call SendData("/CMSG " & stxtbuffer & "")
        'Borrar cartel
        ElseIf stxtbuffer = "" Then
            Call SendData(";" & " ")
        'Say
        ElseIf ModoHabla = h.cNormal Then
            Call SendData(";" & stxtbuffer)
        'DENUNCIAS
        ElseIf ModoHabla = h.cDenuncia Then
            Call SendData("/DENUNCIAR " & stxtbuffer & "")
            'PARTY
        ElseIf ModoHabla = h.cParty Then
            Call SendData("/PMSG " & stxtbuffer & "")
            End If
        stxtbuffer = ""
       SendTxt.Text = ""
       KeyCode = 0
        SendTxt.Visible = False
End If
End Sub

Private Sub SendCMSTXT_KeyUp(KeyCode As Integer, Shift As Integer)
    'Send text
    If KeyCode = vbKeyReturn Then
        'Say
        If stxtbuffercmsg <> "" Then
            Call SendData("/CMSG " & stxtbuffercmsg)
        End If

        stxtbuffercmsg = ""
        SendCMSTXT.Text = ""
        KeyCode = 0
        Me.SendCMSTXT.Visible = False
    End If
End Sub


Private Sub SendCMSTXT_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack) And _
       Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then _
        KeyAscii = 0
End Sub


Private Sub SendCMSTXT_Change()
    If Len(SendCMSTXT.Text) > 160 Then
        stxtbuffercmsg = "Soy un cheater, avisenle a un GM"
    Else
        stxtbuffercmsg = SendCMSTXT.Text
    End If
End Sub
Private Sub Socket1_Connect()
    Dim ServerIp As String
    Dim Temporal1 As Long
    Dim Temporal As Long
   
   
    ServerIp = Socket1.PeerAddress
    Temporal = InStr(1, ServerIp, ".")
    Temporal1 = ((mid$(ServerIp, 1, Temporal - 1) Xor &H65) And &H7F) * 16777216
    ServerIp = mid$(ServerIp, Temporal + 1, Len(ServerIp))
    Temporal = InStr(1, ServerIp, ".")
    Temporal1 = Temporal1 + (mid$(ServerIp, 1, Temporal - 1) Xor &HF6) * 65536
    ServerIp = mid$(ServerIp, Temporal + 1, Len(ServerIp))
    Temporal = InStr(1, ServerIp, ".")
    Temporal1 = Temporal1 + (mid$(ServerIp, 1, Temporal - 1) Xor &H4B) * 256
    ServerIp = mid$(ServerIp, Temporal + 1, Len(ServerIp)) Xor &H42
    MixedKey = (Temporal1 + ServerIp)
   
    Second.Enabled = True
   
    'If frmCrearPersonaje.Visible Then
    If EstadoLogin = E_MODO.CrearNuevoPj Then
        Call Login
 
    ElseIf EstadoLogin = E_MODO.Normal Then
        Call Login
 
    ElseIf EstadoLogin = E_MODO.Dados Then
   frmCrearPersonaje.Show vbModal
   
   ElseIf EstadoLogin = E_MODO.CrearAccount Then
   frmCrearAccount.Show vbModal
   
   ElseIf EstadoLogin = E_MODO.LoginAccount Then
   Call Login
   
   ElseIf EstadoLogin = E_MODO.BorrarPj Then
   Call Login

    End If
End Sub

Private Sub Socket1_Disconnect()
    Dim I As Long
    
    
    Second.Enabled = False
    logged = False
    Connected = False
    
    Socket1.Cleanup
    
    frmConnect.MousePointer = vbNormal
    
    If frmCrearPersonaje.Visible = True Then frmConnect.Visible = True
    
    On Local Error Resume Next
    For I = 0 To Forms.Count - 1
        If Forms(I).Name <> Me.Name And Forms(I).Name <> frmConnect.Name Then
            Unload Forms(I)
        End If
    Next I
    On Local Error GoTo 0
    
    frmMain.Visible = False

    pausa = False
    UserMeditar = False
    
#If SegudidadAlkon Then
    LOGGING = False
    LOGSTRING = False
    LastPressed = 0
    LastMouse = False
    LastAmount = 0
#End If

    UserClase = ""
    UserSexo = ""
    UserRaza = ""
    UserEmail = ""
    
    For I = 1 To NUMSKILLS
        UserSkills(I) = 0
    Next I

    For I = 1 To NUMATRIBUTOS
        UserAtributos(I) = 0
    Next I

    SkillPoints = 0
    Alocados = 0

End Sub

Private Sub Socket1_LastError(ErrorCode As Integer, ErrorString As String, Response As Integer)
    '*********************************************
    'Handle socket errors
    '*********************************************
    If ErrorCode = 24036 Then
    frmMensaje.Show
        frmMensaje.MSG.Caption = "Por favor espere, intentando completar conexion."
        Exit Sub
    End If
    frmMensaje.Show
    frmMensaje.MSG.Caption = "Conexión rechazada por el Servidor"
    frmConnect.MousePointer = 1
    Response = 0

    Second.Enabled = False

    frmMain.Socket1.Disconnect
    
    If frmConnect.Visible Then
        frmConnect.Visible = False
    End If

    If Not frmCrearPersonaje.Visible Then
        If Not frmCambiarPass.Visible Then
            frmConnect.Show
        End If
    Else
        frmCrearPersonaje.MousePointer = 0
    End If
End Sub

Private Sub Socket1_Read(DataLength As Integer, IsUrgent As Integer)
    Dim LoopC As Integer

    Dim RD As String
    Dim rBuffer(1 To 500) As String
    Static TempString As String

    Dim CR As Integer
    Dim tChar As String
    Dim sChar As Integer
    Dim Echar As Integer
    Dim aux$
    Dim nfile As Integer
    
    Socket1.Read RD, DataLength
    
    'Check for previous broken data and add to current data
    If TempString <> "" Then
        RD = TempString & RD
        TempString = ""
    End If

    'Check for more than one line
    sChar = 1
    For LoopC = 1 To Len(RD)

        tChar = mid$(RD, LoopC, 1)

        If tChar = ENDC Then
            CR = CR + 1
            Echar = LoopC - sChar
            rBuffer(CR) = mid$(RD, sChar, Echar)
            sChar = LoopC + 1
        End If

    Next LoopC

    'Check for broken line and save for next time
    If Len(RD) - (sChar - 1) <> 0 Then
        TempString = mid$(RD, sChar, Len(RD))
    End If

    'Send buffer to Handle data
    For LoopC = 1 To CR
        'Call LogCustom("HandleData: " & rBuffer(loopc))
        Call HandleData(rBuffer(LoopC))
    Next LoopC
End Sub
Private Sub AbrirMenuViewPort()
#If (ConMenuseConextuales = 1) Then

If tX >= MinXBorder And tY >= MinYBorder And _
    tY <= MaxYBorder And tX <= MaxXBorder Then
    If MapData(tX, tY).charindex > 0 Then
        If charlist(MapData(tX, tY).charindex).invisible = False Then
        
            Dim I As Long
            Dim m As New frmMenuseFashion
            
            Load m
            m.SetCallback Me
            m.SetMenuId 1
            m.ListaInit 2, False
            
            If charlist(MapData(tX, tY).charindex).Nombre <> "" Then
                m.ListaSetItem 0, charlist(MapData(tX, tY).charindex).Nombre, True
            Else
                m.ListaSetItem 0, "<NPC>", True
            End If
            m.ListaSetItem 1, "Comerciar"
            
            m.ListaFin
            m.Show , Me

        End If
    End If
End If

#End If
End Sub

Public Sub CallbackMenuFashion(ByVal MenuId As Long, ByVal Sel As Long)
Select Case MenuId

Case 0 'Inventario
    Select Case Sel
    Case 0
    Case 1
    Case 2 'Tirar
        Call TirarItem
    Case 3 'Usar
        If Not NoPuedeUsar Then
            NoPuedeUsar = True
            Call UsarItem
        End If
    Case 3 'equipar
        Call EquiparItem
    End Select
    
Case 1 'Menu del ViewPort del engine
    Select Case Sel
    Case 0 'Nombre
        SendData "LC" & tX & "," & tY
    Case 1 'Comerciar
        Call SendData("LC" & tX & "," & tY)
        Call SendData("/COMERCIAR")
    End Select
End Select
End Sub


Private Sub Trabajo_Timer()

End Sub

Private Sub TrainingMacro_Timer()

End Sub

'
' -------------------
'    W I N S O C K
' -------------------
'

#If UsarWrench <> 1 Then

Private Sub Winsock1_Close()
    Dim I As Long
    
    Debug.Print "WInsock Close"
    
    LastSecond = 0
    Second.Enabled = False
    logged = False
    Connected = False
    
    If Winsock1.State <> sckClosed Then _
        Winsock1.Close
    
    frmConnect.MousePointer = vbNormal
    
    If frmPasswdSinPadrinos.Visible = True Then frmPasswdSinPadrinos.Visible = False
    frmCrearPersonaje.Visible = False
    frmConnect.Visible = True
    
    On Local Error Resume Next
    For I = 0 To Forms.Count - 1
        If Forms(I).Name <> Me.Name And Forms(I).Name <> frmConnect.Name Then
            Unload Forms(I)
        End If
    Next I
    On Local Error GoTo 0
    
    frmMain.Visible = False

    pausa = False
    UserMeditar = False

    UserClase = ""
    UserSexo = ""
    UserRaza = ""
    UserEmail = ""
    
    For I = 1 To NUMSKILLS
        UserSkills(I) = 0
    Next I

    For I = 1 To NUMATRIBUTOS
        UserAtributos(I) = 0
    Next I

    SkillPoints = 0
    Alocados = 0

    Dialogos.UltimoDialogo = 0
    Dialogos.CantidadDialogos = 0
End Sub

Private Sub Winsock1_Connect()
    Dim ServerIp As String
    Dim Temporal1 As Long
    Dim Temporal As Long
   
    Debug.Print "Winsock Connect"
   
    ServerIp = Winsock1.RemoteHostIP
    Temporal = InStr(1, ServerIp, ".")
    Temporal1 = ((mid$(ServerIp, 1, Temporal - 1) Xor &H65) And &H7F) * 16777216
    ServerIp = mid$(ServerIp, Temporal + 1, Len(ServerIp))
    Temporal = InStr(1, ServerIp, ".")
    Temporal1 = Temporal1 + (mid$(ServerIp, 1, Temporal - 1) Xor &HF6) * 65536
    ServerIp = mid$(ServerIp, Temporal + 1, Len(ServerIp))
    Temporal = InStr(1, ServerIp, ".")
    Temporal1 = Temporal1 + (mid$(ServerIp, 1, Temporal - 1) Xor &H4B) * 256
    ServerIp = mid$(ServerIp, Temporal + 1, Len(ServerIp)) Xor &H42
    MixedKey = (Temporal1 + ServerIp)
   
    Second.Enabled = True
   
    'If frmCrearPersonaje.Visible Then
    If EstadoLogin = E_MODO.CrearNuevoPj Then
        Call Login
 
    ElseIf EstadoLogin = E_MODO.Normal Then
        Call Login
 
    ElseIf EstadoLogin = E_MODO.Dados Then
   frmCrearPersonaje.Show vbModal
   
   ElseIf EstadoLogin = E_MODO.CrearAccount Then
   frmCrearAccount.Show vbModal
   
   ElseIf EstadoLogin = E_MODO.LoginAccount Then
   Call Login
   
   ElseIf EstadoLogin = E_MODO.BorrarPj Then
   Call Login

    End If
End Sub

Private Sub Winsock1_DataArrival(ByVal BytesTotal As Long)
    Dim LoopC As Integer

    Dim RD As String
    Dim rBuffer(1 To 500) As String
    Static TempString As String

    Dim CR As Integer
    Dim tChar As String
    Dim sChar As Integer
    Dim Echar As Integer
    Dim aux$
    Dim nfile As Integer

    Debug.Print "Winsock DataArrival"
    
    'Socket1.Read RD, DataLength
    Winsock1.GetData RD

    'Check for previous broken data and add to current data
    If TempString <> "" Then
        RD = TempString & RD
        TempString = ""
    End If

    'Check for more than one line
    sChar = 1
    For LoopC = 1 To Len(RD)

        tChar = mid$(RD, LoopC, 1)

        If tChar = ENDC Then
            CR = CR + 1
            Echar = LoopC - sChar
            rBuffer(CR) = mid$(RD, sChar, Echar)
            sChar = LoopC + 1
        End If

    Next LoopC

    'Check for broken line and save for next time
    If Len(RD) - (sChar - 1) <> 0 Then
        TempString = mid$(RD, sChar, Len(RD))
    End If

    'Send buffer to Handle data
    For LoopC = 1 To CR
        Call HandleData(rBuffer(LoopC))
    Next LoopC
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    '*********************************************
    'Handle socket errors
    '*********************************************
    frmMensaje.Show
    frmMensaje.MSG.Caption = Description
    frmConnect.MousePointer = 1
    LastSecond = 0
    Second.Enabled = False

    If Winsock1.State <> sckClosed Then _
        Winsock1.Close
    
    If frmOldPersonaje.Visible Then
        frmOldPersonaje.Visible = False
    End If

    If Not frmCrearPersonaje.Visible Then
        If Not frmCambiarPass.Visible Then
            frmConnect.Show
        End If
    Else
        frmCrearPersonaje.MousePointer = 0
    End If
End Sub

#End If


