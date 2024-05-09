Attribute VB_Name = "Mod_Fonts"
Option Explicit

Public Type tFont

    Caracteres(32 To 255) As String

End Type
 
Public Fuentes() As tFont

Public Sub CargarFonts()
    '**************************************************************
    'Autor: Damián Catanzaro (Midraks)
    'Última modificación: 17/03/2011
    'Carga las Fonts.
    '**************************************************************
 
    Dim CantidadFonts As Byte

    Dim i             As Byte, e As Byte
   
    CantidadFonts = GetVar(App.path & "\Init\Fonts.ini", "INIT", "Fonts")
    ReDim Fuentes(CantidadFonts)
   
    For e = 1 To CantidadFonts
        For i = 32 To 255
            Fuentes(e).Caracteres(i) = Left$(GetVar(App.path & "\Init\Fonts.ini", "INIT", "Fuentes(" & e & ").Caracteres(" & i & ")"), 5)
        Next i
    Next e

End Sub
 
Public Sub DibujarTexto(ByVal Texto As String, _
                        y As Integer, _
                        x As Integer, _
                        Optional FontIndex As Byte = 1, _
                        Optional ByVal color As Long)
On Error Resume Next
    '**************************************************************
    'Autor: Damián Catanzaro (Midraks)
    'Última modificación: 17/03/2011
    'Dibujamos las Fonts.
    '**************************************************************
    Dim NLetras     As Integer, i As Integer

    Dim Grh_Index   As Integer

    Dim rgb_list(3) As Long
   
    NLetras = Len(Texto)

    If color <> 0 Then
        rgb_list(0) = color
        rgb_list(1) = color
        rgb_list(2) = color
        rgb_list(3) = color
    Else
        rgb_list(0) = D3DColorXRGB(255, 255, 255)
        rgb_list(1) = D3DColorXRGB(255, 255, 255)
        rgb_list(2) = D3DColorXRGB(255, 255, 255)
        rgb_list(3) = D3DColorXRGB(255, 255, 255)
    End If
   
    For i = 1 To NLetras
        Grh_Index = Fuentes(FontIndex).Caracteres(Asc(mid$(Texto, i, 1)))
        engine.Device_Box_Textured_Render Grh_Index, x, y, GrhData(Grh_Index).pixelWidth, GrhData(Grh_Index).pixelHeight, rgb_list, GrhData(Grh_Index).sX, GrhData(Grh_Index).sY
        x = (x + GrhData(Grh_Index).pixelWidth) - 2
    Next i
 
End Sub



