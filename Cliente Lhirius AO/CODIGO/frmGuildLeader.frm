VERSION 5.00
Begin VB.Form frmGuildLeader 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administración del Clan"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5880
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGuildLeader.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   5880
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command9 
      Caption         =   "Propuestas de alianzas"
      Height          =   495
      Left            =   3000
      MouseIcon       =   "frmGuildLeader.frx":000C
      MousePointer    =   99  'Custom
      TabIndex        =   17
      Top             =   5460
      Width           =   2775
   End
   Begin VB.CommandButton Command8 
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
      Height          =   495
      Left            =   3000
      MouseIcon       =   "frmGuildLeader.frx":015E
      MousePointer    =   99  'Custom
      TabIndex        =   15
      Top             =   5970
      Width           =   2775
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Propuestas de paz"
      Height          =   495
      Left            =   3000
      MouseIcon       =   "frmGuildLeader.frx":02B0
      MousePointer    =   99  'Custom
      TabIndex        =   14
      Top             =   4950
      Width           =   2775
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Editar URL de la web del clan"
      Height          =   495
      Left            =   3000
      MouseIcon       =   "frmGuildLeader.frx":0402
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   4440
      Width           =   2775
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Editar Codex o Descripcion"
      Height          =   495
      Left            =   3000
      MouseIcon       =   "frmGuildLeader.frx":0554
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   3930
      Width           =   2775
   End
   Begin VB.Frame Frame3 
      Caption         =   "Clanes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   2895
      Begin VB.ListBox guildslist 
         Height          =   1425
         ItemData        =   "frmGuildLeader.frx":06A6
         Left            =   120
         List            =   "frmGuildLeader.frx":06A8
         TabIndex        =   11
         Top             =   240
         Width           =   2655
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Detalles"
         Height          =   375
         Left            =   120
         MouseIcon       =   "frmGuildLeader.frx":06AA
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   1800
         Width           =   2655
      End
   End
   Begin VB.Frame txtnews 
      Caption         =   "GuildNews"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   0
      TabIndex        =   6
      Top             =   2280
      Width           =   5775
      Begin VB.CommandButton Command3 
         Caption         =   "Actualizar"
         Height          =   375
         Left            =   120
         MouseIcon       =   "frmGuildLeader.frx":07FC
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   1080
         Width           =   5535
      End
      Begin VB.TextBox txtguildnews 
         Height          =   735
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   240
         Width           =   5535
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Miembros"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   2880
      TabIndex        =   3
      Top             =   0
      Width           =   2895
      Begin VB.CommandButton Command2 
         Caption         =   "Detalles"
         Height          =   375
         Left            =   120
         MouseIcon       =   "frmGuildLeader.frx":094E
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Top             =   1800
         Width           =   2655
      End
      Begin VB.ListBox members 
         Height          =   1425
         ItemData        =   "frmGuildLeader.frx":0AA0
         Left            =   120
         List            =   "frmGuildLeader.frx":0AA2
         TabIndex        =   4
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Solicitudes de ingreso"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   3960
      Width           =   2895
      Begin VB.CommandButton cmdElecciones 
         Caption         =   "Abrir elecciones"
         Height          =   375
         Left            =   120
         MouseIcon       =   "frmGuildLeader.frx":0AA4
         MousePointer    =   99  'Custom
         TabIndex        =   18
         Top             =   1935
         Width           =   2655
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Detalles"
         Height          =   375
         Left            =   120
         MouseIcon       =   "frmGuildLeader.frx":0BF6
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   1170
         Width           =   2655
      End
      Begin VB.ListBox solicitudes 
         Height          =   840
         ItemData        =   "frmGuildLeader.frx":0D48
         Left            =   120
         List            =   "frmGuildLeader.frx":0D4A
         TabIndex        =   1
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Miembros 
         Caption         =   "El clan cuenta con x miembros"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1620
         Width           =   2535
      End
   End
End
Attribute VB_Name = "frmGuildLeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdElecciones_Click()
    Call SendData("ABREELEC")
    Unload Me
End Sub

Private Sub Command1_Click()

frmCharInfo.frmsolicitudes = True
Call SendData("1HRINFO<" & solicitudes.List(solicitudes.ListIndex))

'Unload Me

End Sub

Private Sub Command2_Click()

frmCharInfo.frmmiembros = True
Call SendData("1HRINFO<" & members.List(members.ListIndex))

'Unload Me

End Sub

Private Sub Command3_Click()

Dim k$

k$ = Replace(txtguildnews, vbCrLf, "º")

Call SendData("ACTGNEWS" & k$)

End Sub

Private Sub Command4_Click()
If ReadField(1, guildslist.List(guildslist.ListIndex), Asc("-")) = "CLAN CERRADO" Then Exit Sub
frmGuildBrief.EsLeader = True
Call SendData("CLANDETAILS" & ReadField(1, guildslist.List(guildslist.ListIndex), Asc("-")))

'Unload Me

End Sub

Private Sub Command5_Click()

Call frmGuildDetails.Show(vbModal, frmGuildLeader)

'Unload Me

End Sub

Private Sub Command6_Click()
Call frmGuildURL.Show(vbModeless, frmGuildLeader)
'Unload Me
End Sub

Private Sub Command7_Click()
Call SendData("ENVPROPP")
End Sub
Private Sub Command9_Click()
Call SendData("ENVALPRO")
End Sub


Private Sub Command8_Click()
Unload Me
frmMain.SetFocus
End Sub


Public Sub ParseLeaderInfo(ByVal Data As String)

If Me.Visible Then Exit Sub

Dim r%, T%

r% = Val(ReadField(1, Data, Asc("¬")))

For T% = 1 To r%
    guildslist.AddItem ReadField(1 + T%, Data, Asc("¬"))
Next T%

r% = Val(ReadField(T% + 1, Data, Asc("¬")))
Miembros.Caption = "El clan cuenta con " & r% & " miembros."

Dim k%

For k% = 1 To r%
    members.AddItem ReadField(T% + 1 + k%, Data, Asc("¬"))
Next k%

txtguildnews = Replace(ReadField(T% + k% + 1, Data, Asc("¬")), "º", vbCrLf)

T% = T% + k% + 2

r% = Val(ReadField(T%, Data, Asc("¬")))

For k% = 1 To r%
    solicitudes.AddItem ReadField(T% + k%, Data, Asc("¬"))
Next k%

Me.Show , frmMain

End Sub


Private Sub Form_Deactivate()
'Me.SetFocus
End Sub

Private Sub Form_Load()
If frmOpciones.OptTrans.value = Checked Then Call Aplicar_Transparencia(Me.hWnd, CByte(frmOpciones.Transp.value))
End Sub
