VERSION 5.00
Begin VB.Form frmCanjes 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Sistema de Canjeo"
   ClientHeight    =   5175
   ClientLeft      =   420
   ClientTop       =   315
   ClientWidth     =   4785
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   4785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   600
      Left            =   2800
      ScaleHeight     =   600
      ScaleWidth      =   555
      TabIndex        =   9
      Top             =   790
      Width           =   555
   End
   Begin VB.TextBox lDescripcion 
      BackColor       =   &H00000040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   855
      Left            =   2760
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   3010
      Width           =   1815
   End
   Begin VB.ListBox ListaPremios 
      BackColor       =   &H00000040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   4155
      ItemData        =   "Form1.frx":57E2
      Left            =   240
      List            =   "Form1.frx":57E4
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label lPuntos 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "99999"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   3540
      TabIndex        =   7
      Top             =   240
      Width           =   1005
   End
   Begin VB.Label lCantidad 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   3675
      TabIndex        =   6
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label lAtaque 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   3675
      TabIndex        =   5
      Top             =   1485
      Width           =   855
   End
   Begin VB.Label lDef 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   3675
      TabIndex        =   4
      Top             =   1890
      Width           =   855
   End
   Begin VB.Label lAM 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   3675
      TabIndex        =   3
      Top             =   2310
      Width           =   855
   End
   Begin VB.Label lDM 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   3675
      TabIndex        =   2
      Top             =   2685
      Width           =   855
   End
   Begin VB.Label Requiere 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "99999"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   3600
      TabIndex        =   1
      Top             =   4125
      Width           =   1005
   End
   Begin VB.Image bSalir 
      Height          =   525
      Left            =   3085
      Top             =   4550
      Width           =   1560
   End
   Begin VB.Image bAceptar 
      Height          =   510
      Left            =   120
      Top             =   4550
      Width           =   2460
   End
End
Attribute VB_Name = "frmCanjes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bAceptar_Click()
Call SendData("SPX" & ListaPremios.ListIndex + 1)
End Sub
     
Private Sub ListaPremios_Click()
Call SendData("IPX" & ListaPremios.ListIndex + 1)
End Sub

Private Sub bSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
If frmOpciones.OptTrans.value = Checked Then Call Aplicar_Transparencia(Me.hWnd, CByte(frmOpciones.Transp.value))
Call SendData("IPX" & ListaPremios.ListIndex + 1)
bAceptar.Picture = LoadPicture(App.path & "\Graficos\Principal\Canjear_BcanjearN.jpg")
bSalir.Picture = LoadPicture(App.path & "\Graficos\Principal\Canjear_BsalirN.jpg")
Me.Picture = LoadPicture(App.path & "\Graficos\Principal\Canjear_main.jpg")
End Sub

Private Sub baceptar_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
bAceptar.Picture = LoadPicture(App.path & "\Graficos\Principal\Canjear_BcanjearA.jpg")
End Sub

Private Sub baceptar_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
bAceptar.Picture = LoadPicture(App.path & "\Graficos\Principal\Canjear_BcanjearI.jpg")
End Sub

Private Sub bsalir_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
bSalir.Picture = LoadPicture(App.path & "\Graficos\Principal\Canjear_BsalirA.jpg")
End Sub

Private Sub bsalir_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
bSalir.Picture = LoadPicture(App.path & "\Graficos\Principal\Canjear_BsalirI.jpg")
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
bSalir.Picture = LoadPicture(App.path & "\Graficos\Principal\Canjear_BsalirN.jpg")
bAceptar.Picture = LoadPicture(App.path & "\Graficos\Principal\Canjear_BcanjearN.jpg")
End Sub

