VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.Ocx"
Begin VB.Form frmCargando 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ControlBox      =   0   'False
   Icon            =   "frmCargando.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   240
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label LabelCarga 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Candara"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   435
      Left            =   6000
      TabIndex        =   0
      Top             =   8085
      Width           =   105
   End
   Begin VB.Image imgProgress 
      Height          =   585
      Left            =   2220
      Top             =   8040
      Width           =   7515
   End
End
Attribute VB_Name = "frmCargando"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

'MawenAO frmCargando -www.mawenao.net
'Todos los derechos reservados
'agush
'www.gs-zone.org
Dim Directory As String, bDone As Boolean, dError As Boolean, f As Integer
Private porcentajeActual As Integer
 
Private Const PROGRESS_DELAY = 10
Private Const PROGRESS_DELAY_BACKWARDS = 4
Private Const DEFAULT_PROGRESS_WIDTH = 501
Private Const DEFAULT_STEP_FORWARD = 1
Private Const DEFAULT_STEP_BACKWARDS = -3
 
Public Sub progresoConDelay(ByVal porcentaje As Integer)
 
If porcentaje = porcentajeActual Then Exit Sub
 
Dim step As Integer, stepInterval As Integer, timer As Long, TickCount As Long
 
If (porcentaje > porcentajeActual) Then
    step = DEFAULT_STEP_FORWARD
    stepInterval = PROGRESS_DELAY
Else
    step = DEFAULT_STEP_BACKWARDS
    stepInterval = PROGRESS_DELAY_BACKWARDS
End If
 
Do Until compararPorcentaje(porcentaje, porcentajeActual, step)
    Do Until (timer + stepInterval) <= GetTickCount()
        DoEvents
    Loop
    timer = GetTickCount()
    porcentajeActual = porcentajeActual + step
    Call establecerProgreso(porcentajeActual)
Loop
 
End Sub
 
 
Public Sub establecerProgreso(ByVal nuevoPorcentaje As Integer)
 
If nuevoPorcentaje >= 0 And nuevoPorcentaje <= 100 Then
    imgProgress.Width = DEFAULT_PROGRESS_WIDTH * CLng(nuevoPorcentaje) / 100
ElseIf nuevoPorcentaje > 100 Then
    imgProgress.Width = DEFAULT_PROGRESS_WIDTH
Else
    imgProgress.Width = 0
End If
porcentajeActual = nuevoPorcentaje
 
End Sub
 
Private Function compararPorcentaje(ByVal porcentajeTarget As Integer, ByVal porcentajeAct As Integer, ByVal step As Integer) As Boolean
 
If step = DEFAULT_STEP_FORWARD Then
    compararPorcentaje = (porcentajeAct >= porcentajeTarget)
Else
    compararPorcentaje = (porcentajeAct <= porcentajeTarget)
End If
 
End Function
Private Sub Form_Load()
Me.Picture = LoadPicture(DirGraficos & "cargando.jpg")
Me.Icon = LoadPicture(App.path & "\Graficos\Icono.ico")
imgProgress.Picture = LoadPicture(DirGraficos & "cargando_barra.jpg")
End Sub

Function Analizar()
            On Error Resume Next
           
            Dim iX As Integer
            Dim tX As Integer
            Dim DifX As Integer
           
'LINK1            'Variable que contiene el numero de actualización correcto del servidor
    iX = Inet1.OpenURL("http://lhirius.zxq.net/update/VEREXE.txt") 'Host
    tX = LeerInt(App.path & "\INIT\Update.ini")
    DifX = iX - tX
 
            If Not (DifX = 0) Then
If MsgBox("Tu versión no es la actuál, ¿Deseas ejecutar el actualizador automático?.", vbYesNo, "Tierras Perdidas") = vbYes Then
Call ShellExecute(Me.hWnd, "open", App.path & "/Autoupdate.exe", "", "", 1)
'End
End If
End If
End Function
Private Function LeerInt(ByVal Ruta As String) As Integer
    f = FreeFile
    Open Ruta For Input As f
    LeerInt = Input$(LOF(f), #f)
    Close #f
End Function

Private Sub GuardarInt(ByVal Ruta As String, ByVal Data As Integer)
    f = FreeFile
    Open Ruta For Output As f
    Print #f, Data
    Close #f
End Sub
