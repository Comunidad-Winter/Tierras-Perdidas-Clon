VERSION 5.00
Begin VB.Form frmConnect 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Tierras Perdidas AO"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ClipControls    =   0   'False
   FillColor       =   &H00000040&
   Icon            =   "frmConnect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox PasswordTxt 
      BackColor       =   &H00000080&
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   6045
      PasswordChar    =   "X"
      TabIndex        =   1
      Top             =   3750
      Width           =   2535
   End
   Begin VB.TextBox NameTxt 
      BackColor       =   &H00000080&
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   345
      Left            =   6045
      TabIndex        =   0
      Top             =   3240
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   0
      Left            =   7920
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   1
      Left            =   5520
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Image Image2 
      Height          =   1095
      Left            =   7560
      Top             =   5640
      Width           =   2295
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
        Call SaveGameini
        frmConnect.MousePointer = 1
        frmMain.MousePointer = 1
        prgRun = False
        Call UnloadAllForms
       
End
End If
End Sub

Private Sub Form_Load()
    '[CODE 002]:MatuX
    'EngineRun = False
    '[END]
    
 Dim j
 For Each j In Image1()
    j.Tag = "0"
 Next
 '[CODE]:MatuX
 '
 '  El código para mostrar la versión se genera acá para
 ' evitar que por X razones luego desaparezca, como suele
 ' pasar a veces :)
 '[END]'

End Sub



Private Sub Image1_Click(Index As Integer)


CurServer = 0
IPdelServidor = "127.0.0."
PuertoDelServidor = "7666"


Call Audio.PlayWave(SND_CLICK)

Select Case Index
    Case 0
        
       EstadoLogin = CrearAccount
#If UsarWrench = 1 Then
       If frmMain.Socket1.Connected Then
            frmMain.Socket1.Disconnect
            frmMain.Socket1.Cleanup
            DoEvents
        End If
        frmMain.Socket1.HostAddress = CurServerIp
        frmMain.Socket1.RemotePort = CurServerPort
        frmMain.Socket1.Connect
#Else
        If frmMain.Winsock1.State <> sckClosed Then
            frmMain.Winsock1.Close
            DoEvents
        End If
        frmMain.Winsock1.Connect CurServerIp, CurServerPort
#End If

        
    Case 1
        nombrecuent = NameTxt.Text
        passcuent = PasswordTxt.Text
#If UsarWrench = 1 Then
        If frmMain.Socket1.Connected Then
        frmMain.Socket1.Disconnect
        frmMain.Socket1.Cleanup
        DoEvents
        End If
#Else
        If frmMain.Winsock1.State <> sckClosed Then _
            frmMain.Winsock1.Close
#End If
      '  If frmConnect.MousePointer = 99 Then
      '      Exit Sub
     '   End If
        
        
        'update user info
        nombrecuent = NameTxt.Text
        Dim aux As String
        aux = PasswordTxt.Text
#If SeguridadAlkon Then
        UserPassword = md5.GetMD5String(aux)
        Call md5.MD5Reset
#Else
        UserPassword = aux
#End If
        If CheckUserData(False) = True Then
            EstadoLogin = LoginAccount
            Me.MousePointer = 99
#If UsarWrench = 1 Then
            frmMain.Socket1.HostAddress = CurServerIp
            frmMain.Socket1.RemotePort = CurServerPort
            frmMain.Socket1.Connect
    
#Else
            'If frmMain.Winsock1.State <> sckClosed Then _
               ' frmMain.Winsock1.Close
            frmMain.Winsock1.Connect CurServerIp, CurServerPort
#End If

        End If
        

End Select
Exit Sub


End Sub

Private Sub Image2_Click()
            frmMain.Socket1.HostAddress = CurServerIp
            frmMain.Socket1.RemotePort = CurServerPort
            frmMain.Socket1.Connect
            
            
            frmRecuperar.Visible = True

End Sub
Private Sub NameTxt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call Image1_Click(1)
    End If
End Sub

Private Sub PasswordTxt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call Image1_Click(1)
    End If
End Sub

