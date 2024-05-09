Attribute VB_Name = "Mod_Busquedas"
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'Author: J.A.O (Juan Agu$tín Oliva) / AGUSHH
'mod_Buscados
'    ---
'MODULO FUNCTION ; CAZA DE PERSONAJES ONLINE
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
 
Option Explicit
 
Public Const RECOMPENZA_BUSCADO = 2000000
Public Const MINUTOS_BUSQUEDA = 40
 
'///JAO/// _
Con esto verificamos que no exista ningún usuario buscado, asi evitamos _
buguear todo =P / A ESTA FUNCIÓN NO LE DAMOS USO EN ESTE CÓDIGO.
 
Private Function ExisteUserBuscado(ByVal UserIndex As Integer) As Boolean
If UserList(UserIndex).Buscado > 0 Then
ExisteUserBuscado = True
Else
ExisteUserBuscado = False
End If
End Function
 
Public Sub BuscaUsuarioACazar()
On Error GoTo jao
Dim i As Integer, UserIndex As Integer
Dim UserCazado As Byte
For i = 1 To LastUser
If i < 1 Then Exit Sub
With UserList(i)
If BuscadoExistente = True Then Exit Sub
If Criminal(i) = True And UserList(i).Stats.ELV >= STAT_MAXELV - 11 And UserList(i).Stats.GLD > RECOMPENZA_BUSCADO And MapInfo(.Pos.Map).Pk = True And UserList(i).flags.Muerto = 0 And UserList(i).flags.Privilegios < PlayerType.Consejero Then   ' 1kk JAO
UserCazado = 1
Else
UserCazado = 0
End If
If UltimoIndexBuscado = i Then UserCazado = 0
 
'''''''JAO'''''''
If UserCazado > 0 Then
UserList(i).Buscado = 1
UserList(i).TiempoBuscado = 0
UltimoIndexBuscado = i
BuscadoExistente = True
Call SendData(SendTarget.ToIndex, 0, 0, "||WANTED> " & UCase$(UserList(i).name) & " es el nuevo buscado en el mundo." & FONTTYPE_ROJON)
Else
'Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No hay ningún buscado en este momento..." & FONTTYPE_CENTINELA)
End If
End With
Next i
jao:  'JAO
End Sub
 
'FUNCIÓN DEL CÓDIGO : _
Cada X tiempo se verifica que no exista ningún personaje buscado, si es asi se selecciona uno al azar de todos los personajes onlines existentes. El personaje buscado tiene muchos beneficios, por ejemplo, cada X cantidad de tiempo se le otorga una considerable cantidad de oro. La contra es que al ser asesinado, perderá 1.000.000 monedas de oros (no es una gran perdida ya que ganaría cerca de 3kk). El usuario buscado tiene un límite de 40 minutos con esta caracteristica, una vez concurridos esos 40 minutos, se elije a un nuevo usuario.
'CONDICIONES ... _
Si un usuario asesina al usuario buscado, se elejirá otro usuario al azar. El usuario que asesinó al usuario buscado ganará 1kk, mientras que el usuario buscado los perderá. _
Si el usuario buscado se desconecta, busca a uno nuevo _
El usuario buscado debe estar en un mapa inseguro, no ser newbie, ser criminal y no ser administrador/gm. _
Solo puede existir un personaje buscado .

