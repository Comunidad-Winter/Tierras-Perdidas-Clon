VERSION 5.00
Begin VB.Form frmSpawnList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Invocar"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3300
   ControlBox      =   0   'False
   Icon            =   "frmSpawnList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   3300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Spawn"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   345
      MouseIcon       =   "frmSpawnList.frx":000C
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   2985
      Width           =   1650
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2025
      MouseIcon       =   "frmSpawnList.frx":015E
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   2985
      Width           =   810
   End
   Begin VB.ListBox lstCriaturas 
      Height          =   2400
      Left            =   345
      TabIndex        =   0
      Top             =   480
      Width           =   2490
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Selecciona la criatura:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   645
      TabIndex        =   3
      Top             =   75
      Width           =   1935
   End
End
Attribute VB_Name = "frmSpawnList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Call SendData("SPA" & lstCriaturas.ListIndex + 1)
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Deactivate()
'Me.SetFocus
End Sub

