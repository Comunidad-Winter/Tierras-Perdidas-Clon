VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmOpciones 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4680
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
   Icon            =   "frmOpciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox OptTrans 
      BackColor       =   &H80000012&
      Caption         =   "Transparencia"
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   480
      TabIndex        =   10
      Top             =   3960
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.CheckBox optReflejo 
      BackColor       =   &H00000000&
      Caption         =   "Reflejo en Agua"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   9
      Top             =   3600
      Width           =   2535
   End
   Begin VB.CheckBox Minimap 
      BackColor       =   &H00000000&
      Caption         =   "Ver Minimapa"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   8
      Top             =   3240
      Width           =   2415
   End
   Begin VB.CheckBox Check4 
      BackColor       =   &H00000000&
      Caption         =   "Sonido"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   1320
      Width           =   3015
   End
   Begin VB.CheckBox Check3 
      BackColor       =   &H00000000&
      Caption         =   "Música"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   1680
      Width           =   3735
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00000000&
      Caption         =   "Luz en el Mouse"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   2040
      Width           =   3135
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Cancelar"
      Height          =   255
      Left            =   2760
      TabIndex        =   4
      Top             =   4800
      Width           =   1575
   End
   Begin VB.CheckBox Sombras 
      BackColor       =   &H00000000&
      Caption         =   "Sombras"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   2880
      Width           =   2895
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00000000&
      Caption         =   "Desplegar menú contextual al clickear un personaje"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   2400
      Width           =   3735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar"
      Height          =   255
      Left            =   360
      MouseIcon       =   "frmOpciones.frx":000C
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   4680
      Width           =   1455
   End
   Begin MSComctlLib.Slider Transp 
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   4320
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   450
      _Version        =   393216
      Max             =   250
      SelStart        =   190
      TickStyle       =   3
      Value           =   190
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Opciones"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   180
      Width           =   2775
   End
End
Attribute VB_Name = "frmOpciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Check2_Click()
LuzMouse = Not LuzMouse
    If Check2.value = 0 Then
        Light.Light_Remove (Light.Light_Find(20))

End If
If Check2.value = Checked Then
        Light.Light_Create UserPos.x + frmMain.MouseX \ 32 - frmMain.renderer.ScaleWidth \ 64, UserPos.y + frmMain.MouseY / 32 - frmMain.renderer.ScaleHeight \ 64, 3, 20, 255, 255, 255
    End If
End Sub

Private Sub Check3_Click()
        If frmOpciones.Check3.value = 0 Then
            Musica = False
            Audio.StopMidi
        End If
        If frmOpciones.Check3.value = Checked Then
            Musica = True
            Call Audio.PlayMIDI(CStr(currentMidi) & ".mid")
        End If
End Sub

Private Sub Check4_Click()
        If frmOpciones.Check4.value = 0 Then
            Sound = False
            Call Audio.StopWave
            RainBufferIndex = 0
            'frmMain.IsPlaying = PlayLoop.plNone
        End If
            If frmOpciones.Check4.value = Checked Then
                Sound = True
        End If
End Sub
Private Sub Command2_Click()
Call WriteVar(App.path & "\INIT\UserOptions.ini", "Opciones", "Musica", frmOpciones.Check3.value)
Call WriteVar(App.path & "\INIT\UserOptions.ini", "Opciones", "Sonidos", frmOpciones.Check4.value)
Call WriteVar(App.path & "\INIT\UserOptions.ini", "Opciones", "MenuPJs", frmOpciones.Check1.value)
Call WriteVar(App.path & "\INIT\UserOptions.ini", "Opciones", "Sombras", frmOpciones.Sombras.value)
Call WriteVar(App.path & "\INIT\UserOptions.ini", "Opciones", "LuzMouse", frmOpciones.Check2.value)
Call WriteVar(App.path & "\INIT\UserOptions.ini", "Opciones", "Minimap", frmOpciones.Minimap.value)
Call WriteVar(App.path & "\INIT\UserOptions.ini", "Opciones", "Reflejo", frmOpciones.optReflejo.value)
Call WriteVar(App.path & "\INIT\UserOptions.ini", "Opciones", "Reflejo", frmOpciones.OptTrans.value)
Me.Visible = False
End Sub
Private Sub Command4_Click()
Unload Me
End Sub

Private Sub optConsola_Click()
    DialogosClanes.Activo = False
End Sub

Private Sub optPantalla_Click()
    DialogosClanes.Activo = True
End Sub


Private Sub Form_Load()
Rem Obtenemos valores de los checks - Opciones Generales...
'If frmMain.Visible = True Then
    Dim Activado As Integer
    Activado = Val(GetVar(App.path & "\INIT\UserOptions.ini", "Opciones", "MenuPJs"))
    frmOpciones.Check1.value = Activado
    
    Activado = Val(GetVar(App.path & "\INIT\UserOptions.ini", "Opciones", "Sombras"))
    frmOpciones.Sombras.value = Activado
    
    Activado = Val(GetVar(App.path & "\INIT\UserOptions.ini", "Opciones", "Musica"))
    Musica = Activado
    
    Activado = Val(GetVar(App.path & "\INIT\UserOptions.ini", "Opciones", "Sonidos"))
    Sound = Activado
    
    Activado = Val(GetVar(App.path & "\INIT\UserOptions.ini", "Opciones", "LuzMouse"))
    frmOpciones.Check2.value = Activado
    'If frmOpciones.Check2.value = 1 Then
    'LuzMouse = True
    'End If
    
    Activado = Val(GetVar(App.path & "\INIT\UserOptions.ini", "Opciones", "Minimap"))
    frmOpciones.Minimap.value = Activado
    
    Activado = Val(GetVar(App.path & "\INIT\UserOptions.ini", "Opciones", "Reflejo"))
    frmOpciones.optReflejo.value = Activado
    
    Activado = Val(GetVar(App.path & "\INIT\UserOptions.ini", "Opciones", "Alpha"))
    frmOpciones.OptTrans.value = Activado

'If frmOpciones.Minimap.value = Unchecked Then
'frmMain.Minimap.Visible = False
'End If
'If frmOpciones.Minimap.value = Checked Then
'frmMain.Minimap.Visible = True
'End If
    
'    If frmOpciones.Check2.value = 0 Then
'        Light.Light_Remove (Light.Light_Find(20))

'End If
'If frmOpciones.Check2.value = Checked Then
'        Light.Light_Create UserPos.x + frmMain.MouseX \ 32 - frmMain.renderer.ScaleWidth \ 64, UserPos.y + frmMain.MouseY / 32 - frmMain.renderer.ScaleHeight \ 64, 3, 20, 255, 255, 255
'    End If
    
'End If

End Sub

Private Sub Minimap_Click()
If Minimap.value = Unchecked Then
frmMain.Minimap.Visible = False
End If
End Sub
