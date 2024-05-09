VERSION 5.00
Begin VB.Form frmPanelGm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Panel GM"
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4905
   Icon            =   "frmPanelGm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   4905
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Crear Torneo / Evento"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   120
      TabIndex        =   23
      Top             =   3600
      Width           =   4695
      Begin VB.ListBox Torneo_Tipo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   360
         TabIndex        =   28
         Top             =   240
         Width           =   3735
      End
      Begin VB.ListBox Cupos_Maximos 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1680
         TabIndex        =   27
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox Nivel_Minimo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         TabIndex        =   26
         Text            =   "1"
         Top             =   960
         Width           =   495
      End
      Begin VB.ListBox Evento_Faccionario 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   1680
         Style           =   1  'Checkbox
         TabIndex        =   25
         Top             =   1320
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Iniciar cuenta regresiva para las inscripciones ..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   2040
         Width           =   4215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cupos Màximos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Cliquea antes de empezar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   31
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label3 
         Caption         =   "Nivel Maximo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Faccionario"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   29
         Top             =   1320
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "Ver oro en banco"
      Height          =   315
      Index           =   19
      Left            =   120
      TabIndex        =   22
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "Show SOS"
      Height          =   315
      Index           =   18
      Left            =   3420
      TabIndex        =   21
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "Boveda"
      Height          =   315
      Index           =   17
      Left            =   2340
      TabIndex        =   20
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "Ban X ip"
      Height          =   315
      Index           =   16
      Left            =   1260
      TabIndex        =   19
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "Penas"
      Height          =   315
      Index           =   15
      Left            =   180
      TabIndex        =   18
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "IP 2 NICK"
      Height          =   315
      Index           =   14
      Left            =   1260
      TabIndex        =   17
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "NICK 2 IP"
      Height          =   315
      Index           =   13
      Left            =   180
      TabIndex        =   16
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "UNBAN"
      Height          =   315
      Index           =   12
      Left            =   3420
      TabIndex        =   15
      Top             =   1860
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "CARCEL"
      Height          =   315
      Index           =   11
      Left            =   3420
      TabIndex        =   14
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "SKILLS"
      Height          =   315
      Index           =   10
      Left            =   1260
      TabIndex        =   13
      Top             =   1860
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "INV"
      Height          =   315
      Index           =   9
      Left            =   180
      TabIndex        =   12
      Top             =   1860
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "INFO"
      Height          =   315
      Index           =   8
      Left            =   3420
      TabIndex        =   11
      Top             =   1020
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "N.ENE."
      Height          =   315
      Index           =   7
      Left            =   180
      TabIndex        =   10
      Top             =   1020
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "DONDE"
      Height          =   315
      Index           =   6
      Left            =   3420
      TabIndex        =   9
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "HORA"
      Height          =   315
      Index           =   5
      Left            =   2340
      TabIndex        =   8
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "Guardar comentario"
      Height          =   315
      Index           =   4
      Left            =   180
      TabIndex        =   7
      Top             =   600
      Width           =   2055
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "IRA"
      Height          =   315
      Index           =   3
      Left            =   1260
      TabIndex        =   6
      Top             =   1020
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "SUM"
      Height          =   315
      Index           =   2
      Left            =   2340
      TabIndex        =   5
      Top             =   1020
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "BAN"
      Height          =   315
      Index           =   1
      Left            =   2340
      TabIndex        =   4
      Top             =   1860
      Width           =   975
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "ECHAR"
      Height          =   315
      Index           =   0
      Left            =   2340
      TabIndex        =   3
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmdActualiza 
      Caption         =   "Actualiza"
      Height          =   315
      Left            =   3600
      TabIndex        =   2
      Top             =   120
      Width           =   915
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   3120
      Width           =   4035
   End
   Begin VB.ComboBox cboListaUsus 
      Height          =   315
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   3435
   End
   Begin VB.Line Line1 
      Index           =   5
      X1              =   120
      X2              =   120
      Y1              =   540
      Y2              =   1380
   End
   Begin VB.Line Line1 
      Index           =   4
      X1              =   4440
      X2              =   4440
      Y1              =   540
      Y2              =   960
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   2280
      X2              =   2280
      Y1              =   960
      Y2              =   1380
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   2280
      X2              =   4440
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   120
      X2              =   4440
      Y1              =   540
      Y2              =   540
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   120
      X2              =   2280
      Y1              =   1380
      Y2              =   1380
   End
End
Attribute VB_Name = "frmPanelGm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAccion_Click(Index As Integer)
Dim Ok As Boolean, tmp As String, Tmp2 As String
Dim Nick As String

Nick = cboListaUsus.Text

Select Case Index
Case 0 '/ECHAR nick
    Call SendData("/ECHAR " & Nick)
Case 1 '/ban motivo@nick
    tmp = InputBox("Motivo ?", "")
    If MsgBox("Esta seguro que desea banear al personaje """ & Nick & """ ?", vbYesNo) = vbYes Then
        Call SendData("/BAN " & tmp & "@" & Nick)
    End If
Case 2 '/sum nick
    Call SendData("/SUM " & Nick)
Case 3 '/ira nick
    Call SendData("/IRA " & Nick)
Case 4 '/rem
    tmp = InputBox("Comentario ?", "")
    Call SendData("/REM " & tmp)
Case 5 '/hora
    Call SendData("/HORA")
Case 6 '/donde nick
    Call SendData("/DONDE " & Nick)
Case 7 '/nene
    tmp = InputBox("Mapa ?", "")
    Call SendData("/NENE " & Trim(tmp))
Case 8 '/info nick
    Call SendData("/INFO " & Nick)
Case 9 '/inv nick
    Call SendData("/INV " & cboListaUsus.Text)
Case 10 '/skills nick
    Call SendData("/SKILLS " & Nick)
Case 11 '/carcel minutos nick
    tmp = InputBox("Minutos ? (hasta 30)", "")
    Tmp2 = InputBox("Razon ?", "")
    If MsgBox("Esta seguro que desea encarcelar al personaje """ & Nick & """ ?", vbYesNo) = vbYes Then
        Call SendData("/CARCEL " & Nick & "@" & Tmp2 & "@" & tmp)
    End If
Case 12 '/unban nick
    If MsgBox("Esta seguro que desea removerle el ban al personaje """ & Nick & """ ?", vbYesNo) = vbYes Then
        Call SendData("/UNBAN " & Nick)
    End If
Case 13 '/nick2ip nick
    Call SendData("/NICK2IP " & Nick)
Case 14 '/ip2nick ip
    Call SendData("/IP2NICK " & Nick)
Case 15 '/penas
    Call SendData("/PENAS " & cboListaUsus.Text)
Case 16 'Ban X ip
    tmp = InputBox("Ingrese el motivo del ban", "Ban X IP")
    If MsgBox("Esta seguro que desea banear el (ip o personaje) " & Nick & "Por IP?", vbYesNo) = vbYes Then
        Nick = Replace(Nick, " ", "+")
        Call SendData("/BANIP " & Nick & tmp)
    End If
Case 17 ' MUESTA BOBEDA
    Call SendData("/BOV " & Nick)
Case 18 ' Sos
    Call SendData("/SHOW SOS")
Case 19 ' Balance
    Call SendData("/BAL " & cboListaUsus.Text)
End Select


End Sub

Private Sub cmdActualiza_Click()
Call SendData("LISTUSU")

End Sub

Private Sub cmdCerrar_Click()
Unload Me
End Sub
Private Sub Command6_Click()
If Val(Cupos_Maximos) < 8 Then
MsgBox "Elija un maximo de cupos por torneo.", vbCritical, "Error #5"
Exit Sub
End If


            Torneos.Cupos = Val(Cupos_Maximos)
Torneos.Nivel = Val(Nivel_Minimo.Text)
Select Case Torneo_Tipo
    Case Is = "Deathmacht - Torneo todos contra todos."
        Torneos.TIPO = "DM"
    Case Is = "Simple Fight - Torneo usuario versus usuario."
        Torneos.TIPO = "1V1"
    Case Is = "Partner Fight - Torneo en pareja (2 versus 2)"
        Torneos.TIPO = "2V2"
    Case Is = "Trilpe Trouble - Torneo en parejas de 3 versus 3"
        Torneos.TIPO = "3V3"
    Case Is = "Guilds Event - Torneo de Clanes"
        Torneos.TIPO = "CE"
    Case Else
        Call MsgBox("Error #3 - Elije un tipo de torneo, P.D.: No olvides de clickearlo.", vbInformation, "Error #3")
    Exit Sub
End Select
Call SendData("/TOR " & Torneos.TIPO & "," & Torneos.Cupos & "," & Torneos.Nivel & "," & Torneos.AutoSum & "," & Torneos.m & "," & Torneos.X & "," & Torneos.Y)
End Sub

Private Sub Form_Load()
    Call cmdActualiza_Click
'TORNEOS
Torneo_Tipo.AddItem "Deathmacht - Torneo todos contra todos."
Torneo_Tipo.AddItem "Simple Fight - Torneo usuario versus usuario."
Torneo_Tipo.AddItem "Partner Fight - Torneo en pareja (2 versus 2)"
Torneo_Tipo.AddItem "Trilpe Trouble - Torneo en parejas de 3 versus 3"
Torneo_Tipo.AddItem "Guilds Event - Torneo de Clanes"
Cupos_Maximos.AddItem "8"
Cupos_Maximos.AddItem "16"
Cupos_Maximos.AddItem "32"
Cupos_Maximos.AddItem "64"
'SOS
'Frame3.Caption = "Preguntas de usuarios y respuestas - (" & MensajesNumber & "/500)"
'Dim ocx As Integer
'For ocx = 1 To MensajesNumber
'    List1.AddItem "[" & MensajesSOS(ocx).Autor & "] - " & MensajesSOS(ocx).TIPO
'Next ocx
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload Me
End Sub

