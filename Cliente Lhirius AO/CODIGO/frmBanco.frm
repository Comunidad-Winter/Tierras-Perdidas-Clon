VERSION 5.00
Begin VB.Form frmBanco 
   BorderStyle     =   0  'None
   Caption         =   "Operación bancaria"
   ClientHeight    =   2940
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2475
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmBanco.frx":0000
   MousePointer    =   99  'Custom
   ScaleHeight     =   2940
   ScaleWidth      =   2475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image1 
      Height          =   255
      Left            =   2160
      Top             =   120
      Width           =   255
   End
   Begin VB.Image retirar 
      Height          =   435
      Left            =   315
      Top             =   2160
      Width           =   1830
   End
   Begin VB.Image depositar 
      Height          =   465
      Left            =   315
      Top             =   1670
      Width           =   1815
   End
   Begin VB.Image Boveda 
      Height          =   450
      Left            =   315
      Top             =   1130
      Width           =   1815
   End
   Begin VB.Label lOro 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100.000.000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   540
      Width           =   1695
   End
End
Attribute VB_Name = "frmBanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Boveda_Click()
SendData ("INIBOV")
End Sub

Private Sub Form_Load()
If frmOpciones.OptTrans.value = Checked Then Call Aplicar_Transparencia(Me.hWnd, CByte(frmOpciones.Transp.value))
    lOro.Caption = PonerPuntos(UserGLDBOV)
    Me.Picture = LoadPicture(App.path & "\Graficos\Principal\Banco_Main.jpg")
    Boveda.Picture = LoadPicture(App.path & "\Graficos\Principal\Banco_BBovedaN.jpg")
    retirar.Picture = LoadPicture(App.path & "\Graficos\Principal\Banco_BretirarN.jpg")
    depositar.Picture = LoadPicture(App.path & "\Graficos\Principal\Banco_BdepositarN.jpg")
End Sub

Private Sub Image1_Click()
Unload frmBanco
End Sub

Private Sub retirar_Click()

On Error Resume Next

Dim Reti As Long
Reti = InputBox$("Ingrese la cantidad", "Retirar")

If Reti <= 0 Then
frmMensaje.Show
frmMensaje.MSG.Caption = "Cantidad inválida."
Exit Sub
End If

If Reti > UserGLDBOV Then
frmMensaje.Show
frmMensaje.MSG.Caption = "No tienes esa cantidad. Escríbela nuevamente."
Exit Sub
Else
Call SendData("/RETIRAR " & Reti)
lOro.Caption = PonerPuntos(UserGLDBOV)
End If

End Sub

Private Sub depositar_Click()

On Error Resume Next

Dim Depo As Long

Depo = InputBox$("Ingrese la cantidad", "Depositar")
If Depo < 0 Then
frmMensaje.Show
frmMensaje.MSG.Caption = "Cantidad inválida."
Exit Sub
End If
        
If Depo > UserGLD Then
frmMensaje.Show
frmMensaje.MSG.Caption = "No tienes esa cantidad. Escríbela nuevamente."
Exit Sub
Else
Call SendData("/DEPOSITAR " & Depo)
lOro.Caption = PonerPuntos(UserGLDBOV)
End If

End Sub

Private Sub boveda_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Boveda.Picture = LoadPicture(App.path & "\Graficos\Principal\Banco_BBovedaa.jpg")
End Sub

Private Sub boveda_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Boveda.Picture = LoadPicture(App.path & "\Graficos\Principal\Banco_BBovedai.jpg")
End Sub

Private Sub depositar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    depositar.Picture = LoadPicture(App.path & "\Graficos\Principal\Banco_Bdepositara.jpg")
End Sub

Private Sub depositar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    depositar.Picture = LoadPicture(App.path & "\Graficos\Principal\Banco_Bdepositari.jpg")
End Sub

Private Sub retirar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    retirar.Picture = LoadPicture(App.path & "\Graficos\Principal\Banco_Bretirara.jpg")
End Sub

Private Sub retirar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    retirar.Picture = LoadPicture(App.path & "\Graficos\Principal\Banco_Bretirari.jpg")
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Boveda.Picture = LoadPicture(App.path & "\Graficos\Principal\Banco_BBovedaN.jpg")
    retirar.Picture = LoadPicture(App.path & "\Graficos\Principal\Banco_BretirarN.jpg")
    depositar.Picture = LoadPicture(App.path & "\Graficos\Principal\Banco_BdepositarN.jpg")
End Sub

