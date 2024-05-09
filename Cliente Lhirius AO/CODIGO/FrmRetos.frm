VERSION 5.00
Object = "{1292FDC1-6231-407E-A10D-F419BBFDA432}#3.0#0"; "ButtonXp.ocx"
Begin VB.Form FrmRetos 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Retos"
   ClientHeight    =   5715
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ButtonXP.XPButton XPButton2 
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   4680
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      Caption         =   "Retar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ButtonXP.XPButton XPButton1 
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   5160
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      Caption         =   "Cancelar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ListBox MiInventario 
      BackColor       =   &H00000080&
      ForeColor       =   &H0000FFFF&
      Height          =   2400
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Width           =   2175
   End
   Begin VB.TextBox Frags 
      BackColor       =   &H00000080&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Text            =   "Frags"
      Top             =   1560
      Width           =   2175
   End
   Begin VB.TextBox PtsTorneo 
      BackColor       =   &H00000080&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Text            =   "Puntos de torneo"
      Top             =   960
      Width           =   2175
   End
   Begin VB.TextBox Nick 
      BackColor       =   &H00004080&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "Nick"
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Si no deseas retar por frags o items o puntos de torneo deja el campo de texto correspondiente como está"
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   2400
      TabIndex        =   12
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   $"FrmRetos.frx":0000
      ForeColor       =   &H000000FF&
      Height          =   1575
      Left            =   2520
      TabIndex        =   11
      Top             =   3960
      Width           =   2055
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Para retar por puntos de torneo necesitas por lo menos 30 puntos."
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Index           =   1
      Left            =   2400
      TabIndex        =   10
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Para retar por Frags necesitas tener mas de 10 frags y 5 frags como minimo para Retar"
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Index           =   0
      Left            =   2400
      TabIndex        =   9
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"FrmRetos.frx":00BE
      ForeColor       =   &H00FFFFFF&
      Height          =   2415
      Left            =   2400
      TabIndex        =   8
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Tu inventario"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Frags"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Puntos de torneo"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nick del oponente"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2145
   End
End
Attribute VB_Name = "FrmRetos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
If frmOpciones.OptTrans.value = Checked Then Call Aplicar_Transparencia(Me.hWnd, CByte(frmOpciones.Transp.value))
            Dim ii As Integer
            ii = 1
            Do While ii <= MAX_INVENTORY_SLOTS
                If Inventario.OBJIndex(ii) <> 0 Then
                        FrmRetos.MiInventario.AddItem Inventario.ItemName(ii)
                Else
                        FrmRetos.MiInventario.AddItem "Nada"
                End If
                ii = ii + 1
            Loop
End Sub
Private Sub XPButton1_Click()
Unload Me
End Sub

Private Sub XPButton2_Click()
    'si no es un item de torneo
    If UCase(FrmRetos.MiInventario.Text) <> "DAGA DE HIELO" And UCase(FrmRetos.MiInventario.Text) <> "DAGA ENVENENADA" And UCase(FrmRetos.MiInventario.Text) <> "ESPADA DE LAS ALMAS" _
       And UCase(FrmRetos.MiInventario.Text) <> "ARMADURA HEROICA" And UCase(FrmRetos.MiInventario.Text) <> "ARMADURA DE CAMPEÓN" And UCase(FrmRetos.MiInventario.Text) <> "ESCUDO DE DRAGÓN" _
       And UCase(FrmRetos.MiInventario.Text) <> "ESPADA ARGENTUM" And UCase(FrmRetos.MiInventario.Text) <> "CASCO OSCURO" And UCase(FrmRetos.MiInventario.Text) <> "CORONA" _
       And UCase(FrmRetos.MiInventario.Text) <> "ESPADA DEL BARLOG" And UCase(FrmRetos.MiInventario.Text) <> "DAGA INFERNAL" And UCase(FrmRetos.MiInventario.Text) <> "CETRO DE ARCHIMAGO" _
       And UCase(FrmRetos.MiInventario.Text) <> "ANILLO DE LOS DIOSES" And UCase(FrmRetos.MiInventario.Text) <> "ARMADURA DE DRAGÓN OSCURO (ALTOS)" And UCase(FrmRetos.MiInventario.Text) <> "ARMADURA DE DRAGÓN OSCURO (BAJOS)" And UCase(FrmRetos.MiInventario.Text) <> "ARCO ARGENTUM" _
       And UCase(FrmRetos.MiInventario.Text) <> "TÚNICA DE APOCALIPSIS (ALTOS)" And UCase(FrmRetos.MiInventario.Text) <> "TÚNICA DE APOCALIPSIS (BAJOS)" And UCase(FrmRetos.MiInventario.Text) <> "MANTO ALADO (ALTOS)" And UCase(FrmRetos.MiInventario.Text) <> "MANTO ALADO (BAJOS)" _
       And UCase(FrmRetos.MiInventario.Text) <> "AMULETO DEL LIDER" And UCase(FrmRetos.MiInventario.Text) <> "PENDIENTE DEL SACRIFICIO" And UCase(FrmRetos.MiInventario.Text) <> "INSURRECCIÓN SOMBRIA" And UCase(FrmRetos.MiInventario.Text) <> "PRISIÓN GELIDA" And UCase(FrmRetos.MiInventario.Text) <> "FUEGO DIVINO" And UCase(FrmRetos.MiInventario.Text) <> "INFIERNO" _
       And UCase(FrmRetos.MiInventario.Text) <> "VIENTO HURACANADO" And UCase(FrmRetos.MiInventario.Text) <> "PERGAMINO DE REGRESO" Then
        
        frmMensaje.Show
        frmMensaje.MSG.Caption = "El objeto seleccionado no está en la lista de Items de torneo."
    Exit Sub
    End If
    
    If UCase(FrmRetos.MiInventario.Text) = "Daga Envenenada" Then ' si el obj vale menos que 30 puntos de torneo
        frmMensaje.Show
        frmMensaje.MSG.Caption = "El objeto seleccionado tiene un valor inferior a 30 puntos de torneo."
    Exit Sub
    End If

     If Not Inventario.Equipped(FrmRetos.MiInventario.ListIndex + 1) Then
     Else
    AddtoRichTextBox frmMain.RecTxt, "No podes retar por el item porque lo estas usando.", 2, 51, 223, 1, 1
     Exit Sub
     End If
        Call SendData("PREPARORETO" & "," & FrmRetos.Nick & "," & FrmRetos.PtsTorneo & "," & FrmRetos.Frags & "," & FrmRetos.MiInventario.Text)


End Sub
