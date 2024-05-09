VERSION 5.00
Begin VB.Form frmPeaceProp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ofertas de paz"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4980
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
   Icon            =   "frmPeaceProp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   4980
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command4 
      Caption         =   "Rechazar"
      Height          =   495
      Left            =   3720
      MouseIcon       =   "frmPeaceProp.frx":000C
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   2520
      MouseIcon       =   "frmPeaceProp.frx":015E
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Detalles"
      Height          =   495
      Left            =   1320
      MouseIcon       =   "frmPeaceProp.frx":02B0
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      MouseIcon       =   "frmPeaceProp.frx":0402
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   2280
      Width           =   975
   End
   Begin VB.ListBox lista 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2010
      ItemData        =   "frmPeaceProp.frx":0554
      Left            =   120
      List            =   "frmPeaceProp.frx":0556
      TabIndex        =   0
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "frmPeaceProp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private tipoprop As TIPO_PROPUESTA
Private Enum TIPO_PROPUESTA
    ALIANZA = 1
    PAZ = 2
End Enum



Private Sub Command1_Click()
Unload Me
End Sub

Public Sub ParsePeaceOffers(ByVal s As String)

Dim T%, r%

T% = Val(ReadField(1, s, 44))

For r% = 1 To T%
    Call lista.AddItem(ReadField(r% + 1, s, 44))
Next r%


tipoprop = PAZ

Me.Show vbModeless, frmMain

End Sub

Public Sub ParseAllieOffers(ByVal s As String)

Dim T%, r%

T% = Val(ReadField(1, s, 44))

For r% = 1 To T%
    Call lista.AddItem(ReadField(r% + 1, s, 44))
Next r%

tipoprop = ALIANZA
Me.Show vbModeless, frmMain

End Sub

Private Sub Command2_Click()
'Me.Visible = False
If tipoprop = PAZ Then
    Call SendData("PEACEDET" & lista.List(lista.ListIndex))
Else
    Call SendData("ALLIEDET" & lista.List(lista.ListIndex))
End If
End Sub

Private Sub Command3_Click()
'Me.Visible = False
If tipoprop = PAZ Then
    Call SendData("ACEPPEAT" & lista.List(lista.ListIndex))
Else
    Call SendData("ACEPALIA" & lista.List(lista.ListIndex))
End If
Me.Hide
Unload Me
End Sub

Private Sub Command4_Click()
If tipoprop = PAZ Then
    Call SendData("RECPPEAT" & lista.List(lista.ListIndex))
Else
    Call SendData("RECPALIA" & lista.List(lista.ListIndex))
End If
Me.Hide
Unload Me
End Sub

Private Sub Form_Load()
If frmOpciones.OptTrans.value = Checked Then Call Aplicar_Transparencia(Me.hWnd, CByte(frmOpciones.Transp.value))
End Sub
