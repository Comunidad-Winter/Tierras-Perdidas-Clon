VERSION 5.00
Begin VB.Form frmOpClick 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   2550
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1935
   Icon            =   "frmOpClick.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   1935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "COMERCIAR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "PARTY"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "DUELO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "DARORO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "PRIVADO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "[SALIR]"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "USUARIO"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1935
   End
End
Attribute VB_Name = "frmOpClick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
Unload Me
End If
End Sub
Private Sub Form_Load()
If frmOpciones.OptTrans.value = Checked Then Call Aplicar_Transparencia(Me.hWnd, CByte(frmOpciones.Transp.value))
 If frmOpciones.Check1.value = 0 Then Unload Me
End Sub

Private Sub Label1_Click()
Call SendData("/COMERCIAR")
Unload Me
End Sub

Private Sub Label2_Click()
Call SendData("/PARTY")
Unload Me
End Sub

Private Sub Label3_Click()
Call SendData("/DUELO " & Label5.Caption & "@200000")
Unload Me
End Sub

Private Sub Label4_Click()
Unload Me
End Sub

Private Sub Label6_Click()
Unload Me
End Sub

Private Sub Label7_Click()
Call SendData("\" & Label5.Caption & " " & InputBox("Escribí el mensaje:", "Mensaje"))
Unload Me
End Sub
