VERSION 5.00
Begin VB.Form frmRespuestaGM 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Respuesta GM"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   3240
   Icon            =   "frmRespuestGm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   3240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3015
   End
End
Attribute VB_Name = "frmRespuestaGM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
If frmOpciones.OptTrans.value = Checked Then Call Aplicar_Transparencia(Me.hWnd, CByte(frmOpciones.Transp.value))
End Sub

Private Sub Form_Unload(Cancel As Integer)
TieneParaResponder = True
frmWriteMSG.Show
End Sub

Private Sub Image1_Click()
Unload Me
End Sub

