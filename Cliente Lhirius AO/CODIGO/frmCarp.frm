VERSION 5.00
Begin VB.Form frmCarp 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Carpintero"
   ClientHeight    =   3210
   ClientLeft      =   -60
   ClientTop       =   -105
   ClientWidth     =   4740
   ControlBox      =   0   'False
   Icon            =   "frmCarp.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3210
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstArmas 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Height          =   2175
      Left            =   240
      TabIndex        =   0
      Top             =   244
      Width           =   4215
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   360
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   2640
      Top             =   2520
      Width           =   1695
   End
End
Attribute VB_Name = "frmCarp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Sub Image1_Click()
On Error Resume Next

Call SendData("CNC" & ObjCarpintero(lstArmas.ListIndex))

Unload Me
End Sub

Private Sub Image2_Click()
Unload Me
End Sub
