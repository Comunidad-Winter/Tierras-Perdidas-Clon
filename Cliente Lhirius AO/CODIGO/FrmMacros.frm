VERSION 5.00
Begin VB.Form FrmMacros 
   Caption         =   "Macros"
   ClientHeight    =   6720
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3825
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FrmMacros.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6720
   ScaleWidth      =   3825
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Tecla Numerica / Comando"
      Height          =   5055
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   3735
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   13
         Text            =   "/COMERCIAR"
         Top             =   240
         Width           =   3255
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   12
         Text            =   "/RESUCITAR"
         Top             =   720
         Width           =   3255
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Index           =   2
         Left            =   360
         TabIndex        =   11
         Text            =   "/CURAR"
         Top             =   1200
         Width           =   3255
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Index           =   3
         Left            =   360
         TabIndex        =   10
         Text            =   "/ONLINE"
         Top             =   1680
         Width           =   3255
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Index           =   4
         Left            =   360
         TabIndex        =   9
         Text            =   "/GM"
         Top             =   2160
         Width           =   3255
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Index           =   5
         Left            =   360
         TabIndex        =   8
         Text            =   "/TORNEO"
         Top             =   2640
         Width           =   3255
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Index           =   6
         Left            =   360
         TabIndex        =   7
         Text            =   "/PARTY"
         Top             =   3120
         Width           =   3255
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Index           =   7
         Left            =   360
         TabIndex        =   6
         Text            =   "/EST"
         Top             =   3600
         Width           =   3255
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Index           =   8
         Left            =   360
         TabIndex        =   5
         Text            =   "N/A"
         Top             =   4080
         Width           =   3255
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Index           =   9
         Left            =   360
         TabIndex        =   4
         Text            =   "N/A"
         Top             =   4560
         Width           =   3255
      End
      Begin VB.Label Label1 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   3240
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   3720
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   4200
         Width           =   1335
      End
      Begin VB.Label Label10 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   4680
         Width           =   1335
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Grabar Configuración"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   6120
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   6120
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      Caption         =   "Como funciona:"
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   5160
      Width           =   3735
      Begin VB.Label Label11 
         Caption         =   "Colocar el comando sin / y presionar los números según corresponda."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   24
         Top             =   240
         Width           =   3255
      End
   End
End
Attribute VB_Name = "FrmMacros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
SaveSetting App.EXEName, "textos", "text1(0)", Text1(0).Text
SaveSetting App.EXEName, "textos", "text1(1)", Text1(1).Text
SaveSetting App.EXEName, "textos", "text1(2)", Text1(2).Text
SaveSetting App.EXEName, "textos", "text1(3)", Text1(3).Text
SaveSetting App.EXEName, "textos", "text1(4)", Text1(4).Text
SaveSetting App.EXEName, "textos", "text1(5)", Text1(5).Text
SaveSetting App.EXEName, "textos", "text1(6)", Text1(6).Text
SaveSetting App.EXEName, "textos", "text1(7)", Text1(7).Text
SaveSetting App.EXEName, "textos", "text1(8)", Text1(8).Text
SaveSetting App.EXEName, "textos", "text1(9)", Text1(9).Text
 
MsgBox "Macros Guardados."
Me.Hide
End Sub

Private Sub Command2_Click()
Me.Hide
End Sub

Private Sub Form_Load()
Text1(0).Text = GetSetting(App.EXEName, "textos", "text1(0)", "")
Text1(1).Text = GetSetting(App.EXEName, "textos", "text1(1)", "")
Text1(2).Text = GetSetting(App.EXEName, "textos", "text1(2)", "")
Text1(3).Text = GetSetting(App.EXEName, "textos", "text1(3)", "")
Text1(4).Text = GetSetting(App.EXEName, "textos", "text1(4)", "")
Text1(5).Text = GetSetting(App.EXEName, "textos", "text1(5)", "")
Text1(6).Text = GetSetting(App.EXEName, "textos", "text1(6)", "")
Text1(7).Text = GetSetting(App.EXEName, "textos", "text1(7)", "")
Text1(8).Text = GetSetting(App.EXEName, "textos", "text1(8)", "")
Text1(9).Text = GetSetting(App.EXEName, "textos", "text1(9)", "")
End Sub
 
 
