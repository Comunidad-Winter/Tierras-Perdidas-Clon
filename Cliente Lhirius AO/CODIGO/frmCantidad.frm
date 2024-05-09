VERSION 5.00
Begin VB.Form frmCantidad 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1110
   ClientLeft      =   1680
   ClientTop       =   4455
   ClientWidth     =   2415
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmCantidad.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1110
   ScaleWidth      =   2415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Tirar todo"
      Height          =   255
      Left            =   1320
      TabIndex        =   2
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Tirar"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2145
   End
End
Attribute VB_Name = "frmCantidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Sub Command2_Click()


frmCantidad.Visible = False
If Inventario.SelectedItem <> FLAGORO Then
    SendData "TI" & Inventario.SelectedItem & "," & Inventario.Amount(Inventario.SelectedItem)
Else
    SendData "TI" & Inventario.SelectedItem & "," & UserGLD
End If

End Sub

Private Sub Command1_Click()
frmCantidad.Visible = False
SendData "TI" & Inventario.SelectedItem & "," & frmCantidad.Text1.Text
frmCantidad.Text1.Text = "0"
End Sub

Private Sub Form_Deactivate()
'Unload Me
End Sub

Private Sub Image2_Click()
frmCantidad.Visible = False
SendData "TI" & Inventario.SelectedItem & "," & frmCantidad.Text1.Text
frmCantidad.Text1.Text = "0"
End Sub

Private Sub Image3_Click()
End Sub

Private Sub text1_Change()
On Error GoTo ErrHandler
    If Val(Text1.Text) < 0 Then
        Text1.Text = MAX_INVENTORY_OBJS
    End If
    
    If Val(Text1.Text) > MAX_INVENTORY_OBJS Then
        If Inventario.SelectedItem <> FLAGORO Or Val(Text1.Text) > UserGLD Then
            Text1.Text = "1"
        End If
    End If
    
    Exit Sub
    
ErrHandler:
    'If we got here the user may have pasted (Shift + Insert) a REALLY large number, causing an overflow, so we set amount back to 1
    Text1.Text = "1"
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
If (KeyAscii <> 8) Then
    If (KeyAscii < 48 Or KeyAscii > 57) Then
        KeyAscii = 0
    End If
End If
End Sub
