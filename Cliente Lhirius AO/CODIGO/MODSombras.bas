Attribute VB_Name = "MODSombras"
Public Function GetAngle2Points(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Double
     
'Note: Calcula el ángulo entre dos puntos
   
    GetAngle2Points = GetAngleXY((X2 - X1), (Y2 - Y1))
   
End Function
 
Public Function GetAngleXY(ByVal x As Double, ByVal y As Double) As Double
 
'Note: Calcula el ángulo entre dos puntos
   
Dim dblres              As Double
 
    dblres = 0
   
    If (y <> 0) Then
        dblres = Radianes2Grados(Atn(x / y))
        If (x <= 0 And y < 0) Then
            dblres = dblres + 180
        ElseIf (x > 0 And y < 0) Then
            dblres = dblres + 180
        ElseIf (x < 0 And y > 0) Then
            dblres = dblres + 360
        End If
    Else
        If (x > 0) Then
            dblres = 90
        ElseIf (x < 0) Then
            dblres = 270
        End If
    End If
   
    GetAngleXY = dblres
   
End Function
 
Public Function Grados2Radianes(ByVal Grados As Double) As Double
 
'Note: Convierte grados en radianes
   
    Grados2Radianes = Grados * (3.14159265358979 / 180) ' PI / 180
   
End Function
 
Public Function Radianes2Grados(ByVal Radianes As Double) As Double
 
'Note: Convierte radianes en grados
 
    Radianes2Grados = Radianes * 180 / 3.14159265358979 ' 180 / PI
   
End Function
Function Get_Distance(ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer) As Single
 
' Author: Emanuel Matías 'Dunkan'
' Note: Distancia entre dos puntos.
 
    Get_Distance = (Abs(X2 - X1) + Abs(Y2 - Y1)) * 0.5
   
End Function
