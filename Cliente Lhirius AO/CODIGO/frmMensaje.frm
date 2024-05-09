VERSION 5.00
Begin VB.Form frmMensaje 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Mensaje"
   ClientHeight    =   2415
   ClientLeft      =   -60
   ClientTop       =   -105
   ClientWidth     =   3990
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmMensaje.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2415
   ScaleWidth      =   3990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label MSG 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "TPAO Clon papa!!"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1335
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3615
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   1485
      Top             =   1755
      Width           =   1365
   End
End
Attribute VB_Name = "frmMensaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Sub Form_Deactivate()
Me.SetFocus
End Sub

Private Sub Form_Load()
Me.Picture = LoadPicture(App.path & "\Graficos\Principal\Mensaje.jpg")

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Image1.Picture = Nothing
frmMensaje.Picture = LoadPicture(App.path & "\Graficos\Principal\mensaje.jpg")
End Sub

Private Sub Image1_Click()
Unload Me
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Image1.Picture = LoadPicture(App.path & "\Graficos\Principal\OKMapretado.jpg")
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Image1.Picture = LoadPicture(App.path & "\Graficos\Principal\OKMEncima.jpg")
End Sub

