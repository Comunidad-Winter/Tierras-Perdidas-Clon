VERSION 5.00
Begin VB.Form frmBonificadores 
   BackColor       =   &H0000FFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2205
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5445
   LinkTopic       =   "Form1"
   ScaleHeight     =   2205
   ScaleWidth      =   5445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblBeneficio 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
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
      Height          =   600
      Index           =   1
      Left            =   1070
      TabIndex        =   1
      Top             =   1340
      Width           =   4200
   End
   Begin VB.Label lblBeneficio 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
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
      Height          =   585
      Index           =   0
      Left            =   1070
      TabIndex        =   0
      Top             =   510
      Width           =   4200
   End
   Begin VB.Image Bonificacion 
      Height          =   690
      Index           =   1
      Left            =   240
      Top             =   1300
      Width           =   720
   End
   Begin VB.Image Bonificacion 
      Height          =   690
      Index           =   0
      Left            =   240
      Top             =   460
      Width           =   720
   End
End
Attribute VB_Name = "frmBonificadores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
If frmOpciones.OptTrans.value = Checked Then Call Aplicar_Transparencia(Me.hWnd, CByte(frmOpciones.Transp.value))
Bonificacion(1).Picture = LoadPicture(App.path & "\Graficos\principal\Bonificadores_BArrivaN.jpg")
Me.Picture = LoadPicture(App.path & "\Graficos\Principal\Bonificadores_Main.jpg")
Bonificacion(0).Picture = LoadPicture(App.path & "\Graficos\principal\Bonificadores_BAbajoN.jpg")
End Sub
