VERSION 5.00
Begin VB.Form frmMapa 
   BorderStyle     =   0  'None
   Caption         =   "Mapa del Mundo"
   ClientHeight    =   4935
   ClientLeft      =   3480
   ClientTop       =   2190
   ClientWidth     =   4035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   4035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton command1 
      BackColor       =   &H0000C0C0&
      Caption         =   "X"
      Enabled         =   0   'False
      Height          =   190
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2040
      Width           =   190
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      TabIndex        =   2
      Top             =   360
      Width           =   135
   End
   Begin VB.Label ss 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Afueras de Anvilmar"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   615
      Left            =   2280
      TabIndex        =   0
      Top             =   4080
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00404040&
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  'Solid
      Height          =   105
      Left            =   1800
      Shape           =   3  'Circle
      Top             =   1300
      Visible         =   0   'False
      Width           =   105
   End
End
Attribute VB_Name = "frmMapa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const RGN_AND = 1
Private Const RGN_OR = 2
Private Const RGN_XOR = 3
Private Const RGN_DIFF = 4
Private Const RGN_COPY = 5
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function OffsetRgn Lib "gdi32" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetRegionData Lib "gdi32" (ByVal hRgn As Long, ByVal dwCount As Long, lpRgnData As Any) As Long

Dim bytRegion() As Byte
Dim nBytes As Long

Private Sub MakeRegion(ByRef frm As Form, ByVal TrnsColor As Long)

frm.BorderStyle = 0
Dim ScaleSize As Long
Dim Width, Height As Long
Dim rgnMain As Long
Dim X, Y As Long
Dim rgnPixel As Long
Dim RGBColor As Long
Dim dcMain As Long
Dim bmpMain As Long
ScaleSize = frm.ScaleMode
frm.ScaleMode = 3
Width = frm.ScaleX(frm.Picture.Width, vbHimetric, vbPixels)
Height = frm.ScaleY(frm.Picture.Height, vbHimetric, vbPixels)
frm.Width = Width * Screen.TwipsPerPixelX
frm.Height = Height * Screen.TwipsPerPixelY
rgnMain = CreateRectRgn(0, 0, Width, Height)
dcMain = CreateCompatibleDC(frm.hDC)
bmpMain = SelectObject(dcMain, frm.Picture.handle)
For Y = 0 To Height
    For X = 0 To Width
        RGBColor = GetPixel(dcMain, X, Y)
        If RGBColor = TrnsColor Then
            rgnPixel = CreateRectRgn(X, Y, X + 1, Y + 1)
            CombineRgn rgnMain, rgnMain, rgnPixel, RGN_XOR
            DeleteObject rgnPixel
        End If
    Next X
Next Y
SelectObject dcMain, bmpMain
DeleteDC dcMain
DeleteObject bmpMain
If rgnMain <> 0 Then
 nBytes = GetRegionData(rgnMain, 0, ByVal 0&)
    If nBytes > 0 Then
        ReDim bytRegion(0 To nBytes - 1)
        nBytes = GetRegionData(rgnMain, nBytes, bytRegion(0))
    End If
    SetWindowRgn frm.hWnd, rgnMain, True
    'CenterForm Me
End If
frm.ScaleMode = 3 'ScaleSize
End Sub '''

Public Sub SetPicture(Filename As String, ClrTransparent As Long)

Set Me.Picture = LoadPicture(Filename)
MakeRegion Me, ClrTransparent
End Sub ''
Private Sub Command1_Click()
ss.Caption = ""
Unload Me
End Sub
'Private Sub Form_KeyPress(KeyAscii As Integer)
'If KeyCode = BindKeys(18).KeyCode Then
'ss.Caption = ""
'Unload Me
'End If
'End Sub



Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
If Configuracion.Alpha_Interfaz_Transparencia > 0 Then MakeTransparent Me.hWnd, Configuracion.Alpha_Interfaz_Transparencia
Me.SetPicture App.Path & "\GRAFICOS\Mapa.bmp", 16777215
End Sub
Private Sub Form_Unload(Cancel As Integer)
ss.Caption = ""
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
FormDrag Me
End Sub

Private Sub Label1_Click()
Unload Me
End Sub
