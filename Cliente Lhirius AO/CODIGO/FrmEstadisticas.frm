VERSION 5.00
Begin VB.Form frmEstadisticas 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estadisticas"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6900
   Icon            =   "FrmEstadisticas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   6900
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   60
      MouseIcon       =   "FrmEstadisticas.frx":000C
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   5880
      Width           =   6810
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   22
      Left            =   3600
      TabIndex        =   38
      Top             =   5640
      Width           =   480
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Estadisticas1"
      Height          =   195
      Index           =   5
      Left            =   240
      TabIndex        =   37
      Top             =   5580
      Width           =   2475
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Estadisticas1"
      Height          =   195
      Index           =   4
      Left            =   240
      TabIndex        =   36
      Top             =   5340
      Width           =   2475
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Estadisticas1"
      Height          =   195
      Index           =   3
      Left            =   240
      TabIndex        =   35
      Top             =   5100
      Width           =   2475
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Estadisticas1"
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   34
      Top             =   4860
      Width           =   2475
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Estadisticas1"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   33
      Top             =   4620
      Width           =   2475
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Estadisticas1"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   32
      Top             =   4380
      Width           =   2475
   End
   Begin VB.Shape Shape4 
      Height          =   1815
      Left            =   120
      Top             =   4020
      Width           =   2715
   End
   Begin VB.Shape Shape3 
      Height          =   2055
      Left            =   120
      Top             =   1920
      Width           =   2715
   End
   Begin VB.Shape Shape1 
      Height          =   1755
      Left            =   120
      Top             =   120
      Width           =   2715
   End
   Begin VB.Shape Shape2 
      Height          =   5715
      Left            =   2880
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Estadisticas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   945
      TabIndex        =   31
      Top             =   4080
      Width           =   1080
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   21
      Left            =   3585
      TabIndex        =   30
      Top             =   5400
      Width           =   480
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   20
      Left            =   3585
      TabIndex        =   29
      Top             =   5175
      Width           =   480
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   19
      Left            =   3585
      TabIndex        =   28
      Top             =   4935
      Width           =   480
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   18
      Left            =   3585
      TabIndex        =   27
      Top             =   4695
      Width           =   480
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   17
      Left            =   3585
      TabIndex        =   26
      Top             =   4470
      Width           =   480
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   16
      Left            =   3585
      TabIndex        =   25
      Top             =   4230
      Width           =   480
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   15
      Left            =   3585
      TabIndex        =   24
      Top             =   3990
      Width           =   480
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   14
      Left            =   3585
      TabIndex        =   23
      Top             =   3765
      Width           =   480
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   13
      Left            =   3585
      TabIndex        =   22
      Top             =   3525
      Width           =   480
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reputacion1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   7
      Left            =   285
      TabIndex        =   21
      Top             =   2760
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Faccion"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   1200
      TabIndex        =   20
      Top             =   1965
      Width           =   660
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   12
      Left            =   3585
      TabIndex        =   19
      Top             =   3285
      Width           =   480
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   11
      Left            =   3585
      TabIndex        =   18
      Top             =   3060
      Width           =   480
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   10
      Left            =   3585
      TabIndex        =   17
      Top             =   2820
      Width           =   480
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   9
      Left            =   3585
      TabIndex        =   16
      Top             =   2580
      Width           =   480
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   8
      Left            =   3585
      TabIndex        =   15
      Top             =   2355
      Width           =   480
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   7
      Left            =   3585
      TabIndex        =   14
      Top             =   2115
      Width           =   480
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   6
      Left            =   3585
      TabIndex        =   13
      Top             =   1875
      Width           =   480
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   5
      Left            =   3585
      TabIndex        =   12
      Top             =   1650
      Width           =   480
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   4
      Left            =   3585
      TabIndex        =   11
      Top             =   1410
      Width           =   480
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   3585
      TabIndex        =   10
      Top             =   1170
      Width           =   480
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   3585
      TabIndex        =   9
      Top             =   945
      Width           =   480
   End
   Begin VB.Label Skills 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   3585
      TabIndex        =   8
      Top             =   720
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Skills"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   4650
      TabIndex        =   7
      Top             =   255
      Width           =   465
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   5
      Left            =   300
      TabIndex        =   5
      Top             =   1365
      Width           =   390
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   4
      Left            =   300
      TabIndex        =   4
      Top             =   1155
      Width           =   390
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   300
      TabIndex        =   3
      Top             =   945
      Width           =   390
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   300
      TabIndex        =   2
      Top             =   735
      Width           =   390
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pablo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   300
      TabIndex        =   1
      Top             =   510
      Width           =   390
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Atributos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   1005
      TabIndex        =   0
      Top             =   210
      Width           =   885
   End
End
Attribute VB_Name = "frmEstadisticas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Command1_Click()
Unload Me
End Sub

Public Sub Iniciar_Labels()
'Iniciamos los labels con los valores de los atributos y los skills
Dim I As Integer
For I = 1 To NUMATRIBUTOS
    Atri(I).Caption = AtributosNames(I) & ": " & UserAtributos(I)
Next
For I = 1 To NUMSKILLS
    Skills(I).Caption = SkillsNames(I) & ": " & UserSkills(I)
Next



With UserEstadisticas
    Label6(0).Caption = "Criminales matados: " & .CriminalesMatados
    Label6(1).Caption = "Ciudadanos matados: " & .CiudadanosMatados
    Label6(2).Caption = "Usuarios matados: " & .UsuariosMatados
    Label6(3).Caption = "NPCs matados: " & .NpcsMatados
    Label6(4).Caption = "Clase: " & .Clase
    Label6(5).Caption = "Tiempo restante en carcel: " & .PenaCarcel
End With

If UserCounters.Alineacion = 1 Then
    Label4(7).ForeColor = &H80&
    Label4(7).Caption = "HORDA INFERNAL"
ElseIf UserCounters.Alineacion = 2 Then
    Label4(7).ForeColor = &HC00000
    Label4(7).Caption = "ALIANZA IMPERIAL"
ElseIf UserCounters.Alineacion = 0 Then
    Label4(7).ForeColor = &H404040
    Label4(7).Caption = "NEUTRAL"
End If
End Sub

Private Sub Form_Load()
If frmOpciones.OptTrans.value = Checked Then Call Aplicar_Transparencia(Me.hWnd, CByte(frmOpciones.Transp.value))
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload Me
End Sub
