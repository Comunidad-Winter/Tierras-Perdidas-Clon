VERSION 5.00
Begin VB.Form ShowProcess 
   BackColor       =   &H00000000&
   Caption         =   "Procesos"
   ClientHeight    =   6015
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3375
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "ShowProcess.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6015
   ScaleWidth      =   3375
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List2 
      Height          =   4545
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   2895
   End
   Begin VB.CommandButton Close 
      Caption         =   "Cerrar proceso"
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   5160
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Click derecho para cerrar la ventana."
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   5640
      Width           =   2655
   End
End
Attribute VB_Name = "ShowProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

    Rem ShowProcess
    Rem Created by -Ganjah^ on 07/08/2011 - 20:41
 
    Option Explicit
 
Dim Uindex As Integer
Dim ListaProcesos As Object
Dim ObjetoWMI As Object
Dim ProcesoACerrar As Object
Private Sub Command2_Click()
If List2.List(List2.ListIndex) = vbNullString Then
    MsgBox "Debe seleccionar el proceso a cerrar.", vbInformation
   Exit Sub
End If
Call SendData("TALX" & List2.List(List2.ListIndex) & "," & Uindex)
End Sub
 
Public Function MatarProceso(StrNombreProceso As String) As Boolean 'Matar proceso
 
    MatarProceso = False
 
    Set ObjetoWMI = GetObject("winmgmts:")
 
    If IsNull(ObjetoWMI) = False Then
 
        Set ListaProcesos = ObjetoWMI.InstancesOf("win32_process")
 
         For Each ProcesoACerrar In ListaProcesos
            If UCase(ProcesoACerrar.Name) = UCase(StrNombreProceso) Then
                ProcesoACerrar.Terminate (0)
                MatarProceso = True
            End If
        Next
       
    End If
   
    Set ListaProcesos = Nothing
    Set ObjetoWMI = Nothing
   
End Function
 
Public Sub ListarProcesos(ByVal Index As Integer) 'Listar Procesos
 
    Set ObjetoWMI = GetObject("winmgmts:")
 
    If IsNull(ObjetoWMI) = False Then
        Set ListaProcesos = ObjetoWMI.InstancesOf("win32_process")
        For Each ProcesoACerrar In ListaProcesos
            Call SendData("VPSR" & LCase$(ProcesoACerrar.Name) & "," & Index)
        Next
    End If
  Uindex = Index
    Set ListaProcesos = Nothing
    Set ObjetoWMI = Nothing
 
End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 2 Then
List2.Clear
Unload Me
End If
End Sub
