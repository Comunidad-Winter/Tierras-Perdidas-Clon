VERSION 5.00
Begin VB.Form frmGuildAdm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lista de Clanes Registrados"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4065
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
   Icon            =   "frmGuildAdm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   4065
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   120
      MouseIcon       =   "frmGuildAdm.frx":000C
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Solicitar Ingreso"
      Height          =   375
      Left            =   1080
      MouseIcon       =   "frmGuildAdm.frx":015E
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   2880
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Detalles"
      Height          =   375
      Left            =   2640
      MouseIcon       =   "frmGuildAdm.frx":02B0
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Frame Frame1 
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
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      Begin VB.ListBox GuildsList 
         Height          =   2010
         ItemData        =   "frmGuildAdm.frx":0402
         Left            =   240
         List            =   "frmGuildAdm.frx":0404
         TabIndex        =   1
         Top             =   360
         Width           =   3255
      End
   End
End
Attribute VB_Name = "frmGuildAdm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

If ReadField(1, guildslist.List(guildslist.ListIndex), Asc("-")) = "CLAN CERRADO" Then Exit Sub
Call SendData("CLANDETAILS" & ReadField(1, guildslist.List(guildslist.ListIndex), Asc("-")))

End Sub


Private Sub Command3_Click()
Unload Me
End Sub

Public Sub ParseGuildList(ByVal Rdata As String)

Dim j As Integer, k As Integer
For j = 0 To guildslist.ListCount - 1
    Me.guildslist.RemoveItem 0
Next j
k = CInt(ReadField(1, Rdata, 44))

For j = 1 To k
    guildslist.AddItem ReadField(1 + j, Rdata, 44)
Next j

Me.Show vbModal, frmMain

End Sub

Private Sub Form_Load()
If frmOpciones.OptTrans.value = Checked Then Call Aplicar_Transparencia(Me.hWnd, CByte(frmOpciones.Transp.value))
End Sub
