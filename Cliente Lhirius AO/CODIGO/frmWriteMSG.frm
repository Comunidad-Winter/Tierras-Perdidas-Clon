VERSION 5.00
Begin VB.Form frmWriteMSG 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Formulario de mensaje a administradores"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4935
   ControlBox      =   0   'False
   Icon            =   "frmWriteMSG.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      Caption         =   "Respuesta"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   9
      Top             =   3960
      Width           =   1095
   End
   Begin VB.OptionButton optConsulta 
      Caption         =   "Consulta regular"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Value           =   -1  'True
      Width           =   1635
   End
   Begin VB.OptionButton optConsulta 
      Caption         =   "Bug"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   3360
      TabIndex        =   7
      Top             =   1080
      Width           =   1515
   End
   Begin VB.OptionButton optConsulta 
      Caption         =   "Sugerencia"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   1800
      TabIndex        =   6
      Top             =   1320
      Width           =   1515
   End
   Begin VB.OptionButton optConsulta 
      Caption         =   "Acusación"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   1800
      TabIndex        =   5
      Top             =   960
      Width           =   1515
   End
   Begin VB.OptionButton optConsulta 
      Caption         =   "Descargo"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   1515
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Enviar"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3960
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmWriteMSG.frx":000C
      Top             =   1680
      Width           =   4695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmWriteMSG.frx":003A
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   60
      TabIndex        =   2
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "frmWriteMSG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

If InStr(1, Text1.Text, ",") Then
    MsgBox "Imposible mandar un mensaje GM con el signo ',' (Coma) adentro del mensaje, edita el mensaje y volve a enviarlo.", vbCritical, "Error #85"
    Exit Sub
End If

If Text1.Text = "" Then
    Call MsgBox("Debes escribir tu mensaje", vbOKOnly, "Error")
    Exit Sub
ElseIf DarIndiceElegido = -1 Then
    Call MsgBox("Debes elegir el motivo de tu consulta", vbOKOnly, "Error")
    Exit Sub
Else
    Call SendData("#" & DarIndiceElegido & "," & Text1.Text)
    Debug.Print "Mande SOS"
    Unload Me
End If

End Sub

Private Function DarIndiceElegido() As Integer

Dim I As Integer

For I = 0 To 4
    If optConsulta(I).value = True Then
        DarIndiceElegido = I
        Exit Function
    End If
Next I

DarIndiceElegido = -1

End Function

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
frmRespuestaGM.Show
Me.Hide
End Sub

Private Sub Form_Load()
If frmOpciones.OptTrans.value = Checked Then Call Aplicar_Transparencia(Me.hWnd, CByte(frmOpciones.Transp.value))
If TieneParaResponder = False Then
    Command3.Enabled = False
Else
    Command3.Enabled = True
End If
End Sub

Private Sub Text1_Click()
If Text1.Text = "Si mandas un GM inadecuado serás penado ..." & vbNewLine & "" Then
    Text1.Text = ""
End If
End Sub
