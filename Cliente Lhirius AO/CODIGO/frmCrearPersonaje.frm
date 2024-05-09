VERSION 5.00
Begin VB.Form frmCrearPersonaje 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   Icon            =   "frmCrearPersonaje.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNombre 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   5205
      TabIndex        =   4
      Top             =   450
      Width           =   3615
   End
   Begin VB.ComboBox lstHogar 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":000C
      Left            =   360
      List            =   "frmCrearPersonaje.frx":0013
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   5760
      Width           =   2565
   End
   Begin VB.ComboBox lstRaza 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      ItemData        =   "frmCrearPersonaje.frx":0023
      Left            =   9480
      List            =   "frmCrearPersonaje.frx":0036
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2190
      Width           =   1740
   End
   Begin VB.ComboBox lstGenero 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      ItemData        =   "frmCrearPersonaje.frx":0063
      Left            =   9480
      List            =   "frmCrearPersonaje.frx":006D
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   3210
      Width           =   1740
   End
   Begin VB.ComboBox lstProfesion 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      ItemData        =   "frmCrearPersonaje.frx":0080
      Left            =   9480
      List            =   "frmCrearPersonaje.frx":00B7
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   4200
      Width           =   1740
   End
   Begin VB.Label Modconstitucion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   2430
      TabIndex        =   37
      Top             =   3840
      Width           =   45
   End
   Begin VB.Label modCarisma 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   2430
      TabIndex        =   36
      Top             =   3360
      Width           =   45
   End
   Begin VB.Label modagilidad 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   2430
      TabIndex        =   35
      Top             =   2400
      Width           =   45
   End
   Begin VB.Label modinteligencia 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   2430
      TabIndex        =   34
      Top             =   2880
      Width           =   45
   End
   Begin VB.Label Modfuerza 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   2430
      TabIndex        =   33
      Top             =   2040
      Width           =   45
   End
   Begin VB.Label Skill 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   21
      Left            =   6075
      TabIndex        =   32
      Top             =   7260
      Width           =   90
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   43
      Left            =   5670
      MouseIcon       =   "frmCrearPersonaje.frx":0151
      MousePointer    =   99  'Custom
      Top             =   7320
      Width           =   135
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   42
      Left            =   6645
      MouseIcon       =   "frmCrearPersonaje.frx":02A3
      MousePointer    =   99  'Custom
      Top             =   7275
      Width           =   135
   End
   Begin VB.Label lbFuerza 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Left            =   2040
      TabIndex        =   31
      Top             =   2040
      Width           =   105
   End
   Begin VB.Label lbAgilidad 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Left            =   2040
      TabIndex        =   30
      Top             =   2400
      Width           =   105
   End
   Begin VB.Label lbConstitucion 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Left            =   2040
      TabIndex        =   29
      Top             =   3840
      Width           =   105
   End
   Begin VB.Label lbInteligencia 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Left            =   2040
      TabIndex        =   28
      Top             =   2880
      Width           =   105
   End
   Begin VB.Label lbCarisma 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Left            =   2040
      TabIndex        =   27
      Top             =   3360
      Width           =   105
   End
   Begin VB.Label Skill 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   6075
      TabIndex        =   26
      Top             =   2610
      Width           =   90
   End
   Begin VB.Label Skill 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   6075
      TabIndex        =   25
      Top             =   2400
      Width           =   90
   End
   Begin VB.Label Skill 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   6075
      TabIndex        =   24
      Top             =   2835
      Width           =   90
   End
   Begin VB.Label Skill 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   3
      Left            =   6075
      TabIndex        =   23
      Top             =   3045
      Width           =   90
   End
   Begin VB.Label Skill 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   4
      Left            =   6075
      TabIndex        =   22
      Top             =   3270
      Width           =   90
   End
   Begin VB.Label Skill 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   5
      Left            =   6075
      TabIndex        =   21
      Top             =   3510
      Width           =   90
   End
   Begin VB.Label Skill 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   6
      Left            =   6075
      TabIndex        =   20
      Top             =   3750
      Width           =   90
   End
   Begin VB.Label Skill 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   7
      Left            =   6075
      TabIndex        =   19
      Top             =   3975
      Width           =   90
   End
   Begin VB.Label Skill 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   8
      Left            =   6075
      TabIndex        =   18
      Top             =   4215
      Width           =   90
   End
   Begin VB.Label Skill 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   9
      Left            =   6075
      TabIndex        =   17
      Top             =   4470
      Width           =   90
   End
   Begin VB.Label Skill 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   10
      Left            =   6075
      TabIndex        =   16
      Top             =   4680
      Width           =   90
   End
   Begin VB.Label Skill 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   11
      Left            =   6075
      TabIndex        =   15
      Top             =   4935
      Width           =   90
   End
   Begin VB.Label Skill 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   12
      Left            =   6075
      TabIndex        =   14
      Top             =   5175
      Width           =   90
   End
   Begin VB.Label Skill 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   13
      Left            =   6075
      TabIndex        =   13
      Top             =   5430
      Width           =   90
   End
   Begin VB.Label Skill 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   14
      Left            =   6075
      TabIndex        =   12
      Top             =   5625
      Width           =   90
   End
   Begin VB.Label Skill 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   15
      Left            =   6075
      TabIndex        =   11
      Top             =   5880
      Width           =   90
   End
   Begin VB.Label Skill 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   16
      Left            =   6075
      TabIndex        =   10
      Top             =   6135
      Width           =   90
   End
   Begin VB.Label Skill 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   17
      Left            =   6075
      TabIndex        =   9
      Top             =   6360
      Width           =   90
   End
   Begin VB.Label Skill 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   18
      Left            =   6075
      TabIndex        =   8
      Top             =   6585
      Width           =   90
   End
   Begin VB.Label Skill 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   19
      Left            =   6075
      TabIndex        =   7
      Top             =   6810
      Width           =   90
   End
   Begin VB.Label Skill 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   20
      Left            =   6075
      TabIndex        =   6
      Top             =   7035
      Width           =   90
   End
   Begin VB.Image boton 
      Height          =   690
      Index           =   0
      Left            =   9840
      MouseIcon       =   "frmCrearPersonaje.frx":03F5
      MousePointer    =   99  'Custom
      Top             =   5400
      Width           =   1440
   End
   Begin VB.Image boton 
      Height          =   735
      Index           =   1
      Left            =   8160
      MouseIcon       =   "frmCrearPersonaje.frx":0547
      MousePointer    =   99  'Custom
      Top             =   5400
      Width           =   1485
   End
   Begin VB.Image boton 
      Height          =   1125
      Index           =   2
      Left            =   840
      MouseIcon       =   "frmCrearPersonaje.frx":0699
      MousePointer    =   99  'Custom
      Top             =   4200
      Width           =   1380
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   41
      Left            =   5670
      MouseIcon       =   "frmCrearPersonaje.frx":07EB
      MousePointer    =   99  'Custom
      Top             =   7080
      Width           =   135
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   40
      Left            =   6645
      MouseIcon       =   "frmCrearPersonaje.frx":093D
      MousePointer    =   99  'Custom
      Top             =   7080
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   39
      Left            =   5670
      MouseIcon       =   "frmCrearPersonaje.frx":0A8F
      MousePointer    =   99  'Custom
      Top             =   6855
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   120
      Index           =   38
      Left            =   6645
      MouseIcon       =   "frmCrearPersonaje.frx":0BE1
      MousePointer    =   99  'Custom
      Top             =   6870
      Width           =   135
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   37
      Left            =   5670
      MouseIcon       =   "frmCrearPersonaje.frx":0D33
      MousePointer    =   99  'Custom
      Top             =   6600
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   105
      Index           =   36
      Left            =   6645
      MouseIcon       =   "frmCrearPersonaje.frx":0E85
      MousePointer    =   99  'Custom
      Top             =   6615
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   150
      Index           =   35
      Left            =   5670
      MouseIcon       =   "frmCrearPersonaje.frx":0FD7
      MousePointer    =   99  'Custom
      Top             =   6360
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   34
      Left            =   6645
      MouseIcon       =   "frmCrearPersonaje.frx":1129
      MousePointer    =   99  'Custom
      Top             =   6360
      Width           =   135
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   33
      Left            =   5670
      MouseIcon       =   "frmCrearPersonaje.frx":127B
      MousePointer    =   99  'Custom
      Top             =   6150
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   32
      Left            =   6645
      MouseIcon       =   "frmCrearPersonaje.frx":13CD
      MousePointer    =   99  'Custom
      Top             =   6120
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   150
      Index           =   31
      Left            =   5670
      MouseIcon       =   "frmCrearPersonaje.frx":151F
      MousePointer    =   99  'Custom
      Top             =   5880
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   30
      Left            =   6645
      MouseIcon       =   "frmCrearPersonaje.frx":1671
      MousePointer    =   99  'Custom
      Top             =   5880
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   150
      Index           =   29
      Left            =   5670
      MouseIcon       =   "frmCrearPersonaje.frx":17C3
      MousePointer    =   99  'Custom
      Top             =   5670
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   150
      Index           =   28
      Left            =   6645
      MouseIcon       =   "frmCrearPersonaje.frx":1915
      MousePointer    =   99  'Custom
      Top             =   5640
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   120
      Index           =   26
      Left            =   6645
      MouseIcon       =   "frmCrearPersonaje.frx":1A67
      MousePointer    =   99  'Custom
      Top             =   5400
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   24
      Left            =   6645
      MouseIcon       =   "frmCrearPersonaje.frx":1BB9
      MousePointer    =   99  'Custom
      Top             =   5160
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   22
      Left            =   6645
      MouseIcon       =   "frmCrearPersonaje.frx":1D0B
      MousePointer    =   99  'Custom
      Top             =   4920
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   20
      Left            =   6645
      MouseIcon       =   "frmCrearPersonaje.frx":1E5D
      MousePointer    =   99  'Custom
      Top             =   4800
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   120
      Index           =   18
      Left            =   6645
      MouseIcon       =   "frmCrearPersonaje.frx":1FAF
      MousePointer    =   99  'Custom
      Top             =   4515
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   120
      Index           =   16
      Left            =   6645
      MouseIcon       =   "frmCrearPersonaje.frx":2101
      MousePointer    =   99  'Custom
      Top             =   4245
      Width           =   135
   End
   Begin VB.Image command1 
      Height          =   120
      Index           =   14
      Left            =   6645
      MouseIcon       =   "frmCrearPersonaje.frx":2253
      MousePointer    =   99  'Custom
      Top             =   4050
      Width           =   135
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   12
      Left            =   6645
      MouseIcon       =   "frmCrearPersonaje.frx":23A5
      MousePointer    =   99  'Custom
      Top             =   3795
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   10
      Left            =   6645
      MouseIcon       =   "frmCrearPersonaje.frx":24F7
      MousePointer    =   99  'Custom
      Top             =   3540
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   150
      Index           =   8
      Left            =   6645
      MouseIcon       =   "frmCrearPersonaje.frx":2649
      MousePointer    =   99  'Custom
      Top             =   3315
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   150
      Index           =   6
      Left            =   6645
      MouseIcon       =   "frmCrearPersonaje.frx":279B
      MousePointer    =   99  'Custom
      Top             =   3105
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   4
      Left            =   6645
      MouseIcon       =   "frmCrearPersonaje.frx":28ED
      MousePointer    =   99  'Custom
      Top             =   2835
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   150
      Index           =   2
      Left            =   6645
      MouseIcon       =   "frmCrearPersonaje.frx":2A3F
      MousePointer    =   99  'Custom
      Top             =   2640
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   0
      Left            =   6645
      MouseIcon       =   "frmCrearPersonaje.frx":2B91
      MousePointer    =   99  'Custom
      Top             =   2400
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   1
      Left            =   5685
      MouseIcon       =   "frmCrearPersonaje.frx":2CE3
      MousePointer    =   99  'Custom
      Top             =   2400
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   27
      Left            =   5670
      MouseIcon       =   "frmCrearPersonaje.frx":2E35
      MousePointer    =   99  'Custom
      Top             =   5385
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   25
      Left            =   5670
      MouseIcon       =   "frmCrearPersonaje.frx":2F87
      MousePointer    =   99  'Custom
      Top             =   5190
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   23
      Left            =   5670
      MouseIcon       =   "frmCrearPersonaje.frx":30D9
      MousePointer    =   99  'Custom
      Top             =   4920
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   21
      Left            =   5670
      MouseIcon       =   "frmCrearPersonaje.frx":322B
      MousePointer    =   99  'Custom
      Top             =   4725
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   19
      Left            =   5670
      MouseIcon       =   "frmCrearPersonaje.frx":337D
      MousePointer    =   99  'Custom
      Top             =   4485
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   17
      Left            =   5670
      MouseIcon       =   "frmCrearPersonaje.frx":34CF
      MousePointer    =   99  'Custom
      Top             =   4215
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   15
      Left            =   5670
      MouseIcon       =   "frmCrearPersonaje.frx":3621
      MousePointer    =   99  'Custom
      Top             =   4020
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   13
      Left            =   5670
      MouseIcon       =   "frmCrearPersonaje.frx":3773
      MousePointer    =   99  'Custom
      Top             =   3780
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   11
      Left            =   5670
      MouseIcon       =   "frmCrearPersonaje.frx":38C5
      MousePointer    =   99  'Custom
      Top             =   3540
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   9
      Left            =   5685
      MouseIcon       =   "frmCrearPersonaje.frx":3A17
      MousePointer    =   99  'Custom
      Top             =   3315
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   7
      Left            =   5685
      MouseIcon       =   "frmCrearPersonaje.frx":3B69
      MousePointer    =   99  'Custom
      Top             =   3120
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   5
      Left            =   5685
      MouseIcon       =   "frmCrearPersonaje.frx":3CBB
      MousePointer    =   99  'Custom
      Top             =   2835
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   3
      Left            =   5685
      MouseIcon       =   "frmCrearPersonaje.frx":3E0D
      MousePointer    =   99  'Custom
      Top             =   2625
      Width           =   150
   End
   Begin VB.Label Puntos 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6120
      TabIndex        =   5
      Top             =   7920
      Width           =   270
   End
End
Attribute VB_Name = "frmCrearPersonaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public SkillPoints As Byte

Function CheckData() As Boolean
If UserRaza = "" Then
    MsgBox "Seleccione la raza del personaje."
    Exit Function
End If

If UserSexo = "" Then
    MsgBox "Seleccione el sexo del personaje."
    Exit Function
End If

If UserClase = "" Then
    MsgBox "Seleccione la clase del personaje."
    Exit Function
End If

If UserHogar = "" Then
    MsgBox "Seleccione el hogar del personaje."
    Exit Function
End If

If SkillPoints > 0 Then
    MsgBox "Asigne los skillpoints del personaje."
    Exit Function
End If

Dim i As Integer
For i = 1 To NUMATRIBUTOS
    If UserAtributos(i) = 0 Then
        MsgBox "Los atributos del personaje son invalidos."
        Exit Function
    End If
Next i

CheckData = True


End Function

Private Sub boton_Click(Index As Integer)

Call Audio.PlayWave(SND_CLICK)

Select Case Index
    Case 0
        
        Dim i As Integer
        Dim k As Object
        i = 1
        For Each k In Skill
            UserSkills(i) = k.Caption
            i = i + 1
        Next
        
        UserName = txtNombre.Text
        
        If Right$(UserName, 1) = " " Then
                UserName = RTrim$(UserName)
                MsgBox "Nombre invalido, se han removido los espacios al final del nombre"
        End If
        
        UserRaza = lstRaza.List(lstRaza.ListIndex)
        UserSexo = lstGenero.List(lstGenero.ListIndex)
        UserClase = lstProfesion.List(lstProfesion.ListIndex)
        
        UserAtributos(1) = Val(lbFuerza.Caption)
        UserAtributos(2) = Val(lbInteligencia.Caption)
        UserAtributos(3) = Val(lbAgilidad.Caption)
        UserAtributos(4) = Val(lbCarisma.Caption)
        UserAtributos(5) = Val(lbConstitucion.Caption)
        
        UserHogar = lstHogar.List(lstHogar.ListIndex)
        
        'Barrin 3/10/03
        If CheckData() Then
            frmMain.Socket1.HostName = CurServerIp
    frmMain.Socket1.RemotePort = CurServerPort
    
    Me.MousePointer = 11
    EstadoLogin = CrearNuevoPj

    If Not frmMain.Socket1.Connected Then
        
        frmMensaje.Show
            frmMensaje.MSG.Caption = "Error: Se ha perdido la conexion con el server."
        Unload Me
    Else
        Call Login
    End If
        End If
        
    Case 1

            Call Audio.PlayMIDI("2.mid")
      
        
        Me.Visible = False
        
        
    Case 2
        Call Audio.PlayWave(SND_DICE)
        Call TirarDados
      
End Select


End Sub


Function RandomNumber(ByVal LowerBound As Variant, ByVal UpperBound As Variant) As Single

Randomize timer

RandomNumber = (UpperBound - LowerBound + 1) * Rnd + LowerBound
If RandomNumber > UpperBound Then RandomNumber = UpperBound

End Function


Private Sub TirarDados()
'lbFuerza.Caption = CInt(RandomNumber(1, 6) + RandomNumber(1, 6) + RandomNumber(1, 6))
'lbInteligencia.Caption = CInt(RandomNumber(1, 6) + RandomNumber(1, 6) + RandomNumber(1, 6))
'lbAgilidad.Caption = CInt(RandomNumber(1, 6) + RandomNumber(1, 6) + RandomNumber(1, 6))
'lbCarisma.Caption = CInt(RandomNumber(1, 6) + RandomNumber(1, 6) + RandomNumber(1, 6))
'lbConstitucion.Caption = CInt(RandomNumber(1, 6) + RandomNumber(1, 6) + RandomNumber(1, 6))

#If UsarWrench = 1 Then
    If frmMain.Socket1.Connected Then
#Else
    If frmMain.Winsock1.State = sckConnected Then
#End If
        Call SendData("TIRDAD")
    End If

End Sub

Private Sub Command1_Click(Index As Integer)
Call Audio.PlayWave(SND_CLICK)

Dim indice
If Index Mod 2 = 0 Then
    If SkillPoints > 0 Then
        indice = Index \ 2
        Skill(indice).Caption = Val(Skill(indice).Caption) + 1
        SkillPoints = SkillPoints - 1
    End If
Else
    If SkillPoints < 10 Then
        
        indice = Index \ 2
        If Val(Skill(indice).Caption) > 0 Then
            Skill(indice).Caption = Val(Skill(indice).Caption) - 1
            SkillPoints = SkillPoints + 1
        End If
    End If
End If

Puntos.Caption = SkillPoints
End Sub

Private Sub Form_Load()
SkillPoints = 10
Puntos.Caption = SkillPoints
Me.Picture = LoadPicture(App.path & "\Graficos\Principal\CP-Interface.jpg")
'imgHogar.Picture = LoadPicture(App.path & "\graficos\CP-Ullathorpe.jpg")

Dim i As Integer
lstProfesion.Clear
For i = LBound(ListaClases) To UBound(ListaClases)
    lstProfesion.AddItem ListaClases(i)
Next i

lstProfesion.ListIndex = 0
lstHogar.ListIndex = 0
'Image1.Picture = LoadPicture(App.path & "\graficos\" & lstProfesion.Text & ".jpg")
Call TirarDados
End Sub

Private Sub lstRaza_Click()
Select Case (lstRaza.List(lstRaza.ListIndex))
    Case Is = "Humano"
        Modfuerza.Caption = "+ 1"
        modagilidad.Caption = "+ 1"
        modinteligencia.Caption = "+ 0"
        modCarisma.Caption = "+ 2"
        Modconstitucion.Caption = "+ 2"
    Case Is = "Elfo"
        Modfuerza.Caption = "- 1"
        modagilidad.Caption = "+ 2"
        modinteligencia.Caption = "+ 2"
        modCarisma.Caption = "+ 1"
        Modconstitucion.Caption = "+ 0"
    Case Is = "Elfo Oscuro"
        Modfuerza.Caption = "+ 2"
        modagilidad.Caption = "+ 0"
        modinteligencia.Caption = "+ 1"
        modCarisma.Caption = "+ 0"
        Modconstitucion.Caption = "+ 1"
    Case Is = "Enano"
        Modfuerza.Caption = "+ 3"
        modagilidad.Caption = "- 2"
        modinteligencia.Caption = "- 5"
        modCarisma.Caption = "- 1"
        Modconstitucion.Caption = "+ 3"
    Case Is = "Gnomo"
        Modfuerza.Caption = "- 3"
        modagilidad.Caption = "+ 3"
        modinteligencia.Caption = "+ 3"
        modCarisma.Caption = "+ 0"
        Modconstitucion.Caption = "- 1"
    End Select
End Sub

Private Sub txtNombre_Change()
txtNombre.Text = LTrim(txtNombre.Text)
End Sub

Private Sub txtNombre_GotFocus()
MsgBox "Sea cuidadoso al seleccionar el nombre de su personaje, Argentum es un juego de rol, un mundo magico y fantastico, si selecciona un nombre obsceno o con connotación politica los administradores borrarán su personaje y no habrá ninguna posibilidad de recuperarlo."
End Sub
