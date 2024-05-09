VERSION 5.00
Begin VB.Form frmPres 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8985
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   Icon            =   "frmPres.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8985
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1125
      Top             =   1230
   End
End
Attribute VB_Name = "frmPres"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim puedo As Boolean

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then If puedo Then Unload Me
End Sub

Private Sub Form_Load()
        Me.Picture = LoadPicture(App.path & "\Graficos\pres" & RandomNumber(1, 5) & ".jpg")
    puedo = False
End Sub

Private Sub Timer1_Timer()
Static ticks As Long

ticks = ticks + 1
If ticks = 1 Then
puedo = True

ElseIf ticks = 2 Then
    'Me.Picture = LoadPicture(App.path & "\Graficos\pres" & RandomNumber(1, 5) & ".jpg")
Else
    Unload Me
End If

End Sub
