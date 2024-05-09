Attribute VB_Name = "ModFacciones"
'Recompensas stylo tpao codigo echo por ZaikO exclusivamente para gs-zone.
 
 
Option Explicit
 
Public ArmaduraImperial1 As Integer 'Primer jerarquia
Public ArmaduraImperial2 As Integer 'Segunda jerarqu�a
Public ArmaduraImperial3 As Integer 'Enanos
Public TunicaMagoImperial As Integer 'Magos
Public TunicaMagoImperialEnanos As Integer 'Magos
 
Public VestimentaImperialHumano As Integer
Public VestimentaImperialEnano As Integer
Public TunicaConspicuaHumano As Integer
Public TunicaConspicuaEnano As Integer
Public ArmaduraNobilisimaHumano As Integer
Public ArmaduraNobilisimaEnano As Integer
Public ArmaduraGranSacerdote As Integer
 
Public VestimentaLegionHumano As Integer
Public VestimentaLegionEnano As Integer
Public TunicaLobregaHumano As Integer
Public TunicaLobregaEnano As Integer
Public TunicaEgregiaHumano As Integer
Public TunicaEgregiaEnano As Integer
Public SacerdoteDemoniaco As Integer
 
Public ArmaduraCaos1 As Integer
Public TunicaMagoCaos As Integer
Public TunicaMagoCaosEnanos As Integer
Public ArmaduraCaos2 As Integer
Public ArmaduraCaos3 As Integer
 
Public Const ExpAlUnirse As Long = 50000
Public Const ExpX100 As Integer = 5000
 
 
Public Sub EnlistarArmadaReal(ByVal UserIndex As Integer)
 
If UserList(UserIndex).Faccion.ArmadaReal = 1 Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "�" & "Ya perteneces a las tropas reales!!! Ve a combatir criminales!!!" & "�" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
    Exit Sub
End If
 
If UserList(UserIndex).Faccion.FuerzasCaos = 1 Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "�" & "Maldito insolente!!! vete de aqui seguidor de las sombras!!!" & "�" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
    Exit Sub
End If
 
If Criminal(UserIndex) Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "�" & "No se permiten criminales en el ejercito imperial!!!" & "�" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
    Exit Sub
End If
 
If UserList(UserIndex).Faccion.CriminalesMatados < 100 Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "�" & "Para unirte a nuestras fuerzas debes matar al menos 30 criminales, solo has matado " & UserList(UserIndex).Faccion.CriminalesMatados & "�" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
    Exit Sub
End If
 
If UserList(UserIndex).Stats.ELV < 25 Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "�" & "Para unirte a nuestras fuerzas debes ser al menos de nivel 25!!!" & "�" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
    Exit Sub
End If
 
If UserList(UserIndex).Faccion.CiudadanosMatados > 0 Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "�" & "Has asesinado gente inocente, no aceptamos asesinos en las tropas reales!" & "�" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
    Exit Sub
End If
 
If UserList(UserIndex).Faccion.Reenlistadas > 4 Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "�" & "Has sido expulsado de las fuerzas reales demasiadas veces!" & "�" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
    Exit Sub
End If
 
UserList(UserIndex).Faccion.ArmadaReal = 1
UserList(UserIndex).Faccion.Reenlistadas = UserList(UserIndex).Faccion.Reenlistadas + 1
 
UserList(UserIndex).Faccion.RecompensasReal = UserList(UserIndex).Faccion.CriminalesMatados \ 100
 
Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "�" & "���Bienvenido a al Ejercito Imperial!!!, aqui tienes tus vestimentas. Por cada centena de criminales que acabes te dar� un recompensa, buena suerte soldado!" & "�" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
 
If UserList(UserIndex).Faccion.RecibioArmaduraReal = 0 Then
    Dim MiObj As Obj
    Dim MiObj2 As Obj
    MiObj.Amount = 20
    MiObj2.Amount = 10
   
   
   
   
'Public VestimentaImperialHumano As Integer
'Public VestimentaImperialEnano As Integer
'Public TunicaConspicuaHumano As Integer
'Public TunicaConspicuaEnano As Integer
'Public ArmaduraNobilisimaHumano As Integer
'Public ArmaduraNobilisimaEnano As Integer
'Public ArmaduraGranSacerdote As Integer
 
'Public VestimentaLegionHumano As Integer
'Public VestimentaLegionEnano As Integer
'Public TunicaLobregaHumano As Integer
'Public TunicaLobregaEnano As Integer
'Public TunicaEgregiaHumano As Integer
'Public TunicaEgregiaEnano As Integer
'Public SacerdoteDemoniaco As Integer
'
   
       
    If UCase$(UserList(UserIndex).Raza) = "ENANO" Or UCase$(UserList(UserIndex).Raza) = "GNOMO" Then
        MiObj.ObjIndex = VestimentaImperialEnano
        Select Case UCase$(UserList(UserIndex).Clase)
            Case "MAGO"
                MiObj2.ObjIndex = TunicaConspicuaEnano
            Case Else
                MiObj2.ObjIndex = ArmaduraNobilisimaEnano
        End Select
    Else
        MiObj.ObjIndex = VestimentaImperialHumano
        Select Case UCase$(UserList(UserIndex).Clase)
            Case "MAGO"
                MiObj2.ObjIndex = TunicaConspicuaHumano
            Case "CLERIGO", "DRUIDA", "BARDO"
                MiObj2.ObjIndex = ArmaduraGranSacerdote
            Case Else
                MiObj2.ObjIndex = ArmaduraNobilisimaHumano
        End Select
    End If
   
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If
    If Not MeterItemEnInventario(UserIndex, MiObj2) Then
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj2)
    End If
   
    UserList(UserIndex).Faccion.RecibioArmaduraReal = 1
End If
 
If UserList(UserIndex).Faccion.RecibioExpInicialReal = 0 Then
    UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + ExpAlUnirse
    If UserList(UserIndex).Stats.Exp > MAXEXP Then _
        UserList(UserIndex).Stats.Exp = MAXEXP
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has ganado " & ExpAlUnirse & " puntos de experiencia." & FONTTYPE_DAREXP)
    UserList(UserIndex).Faccion.RecibioExpInicialReal = 1
    Call CheckUserLevel(UserIndex)
End If
 
 
Call LogEjercitoReal(UserList(UserIndex).name)
 
End Sub
 
Public Sub RecompensaArmadaReal(ByVal UserIndex As Integer)
 
Dim Sjerar As Obj
Dim Tjerar As Obj
Dim Cjerar As Obj
 
Sjerar.Amount = 1
Sjerar.ObjIndex = 865
 
Tjerar.Amount = 1
Tjerar.ObjIndex = 889
 
Cjerar.Amount = 1
Cjerar.ObjIndex = 866
 
 
 
If UserList(UserIndex).Faccion.CriminalesMatados \ 100 = _
   UserList(UserIndex).Faccion.RecompensasReal Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "�" & "Ya has recibido tu recompensa, mata 100 crinales mas para recibir la proxima!!!" & "�" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
End If
 
If UserList(UserIndex).flags.PJerarquia = 1 Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "�" & "Aqui tienes tu recompensa noble guerrero!!!" & "�" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
    UserList(UserIndex).Faccion.RecompensasReal = UserList(UserIndex).Faccion.RecompensasReal + 1
    Call CheckUserLevel(UserIndex)
If Not MeterItemEnInventario(UserIndex, Sjerar) Then
    Call TirarItemAlPiso(UserList(UserIndex).Pos, Sjerar)
    End If
    End If
 
If UserList(UserIndex).flags.SJerarquia = 1 Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "�" & "Aqui tienes tu recompensa noble guerrero!!!" & "�" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
    UserList(UserIndex).Faccion.RecompensasReal = UserList(UserIndex).Faccion.RecompensasReal + 1
    Call CheckUserLevel(UserIndex)
If Not MeterItemEnInventario(UserIndex, Tjerar) Then
    Call TirarItemAlPiso(UserList(UserIndex).Pos, Tjerar)
    End If
    End If
 
If UserList(UserIndex).flags.TJerarquia = 1 Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "�" & "Aqui tienes tu recompensa noble guerrero!!!" & "�" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
    UserList(UserIndex).Faccion.RecompensasReal = UserList(UserIndex).Faccion.RecompensasReal + 1
    Call CheckUserLevel(UserIndex)
If Not MeterItemEnInventario(UserIndex, Cjerar) Then
    Call TirarItemAlPiso(UserList(UserIndex).Pos, Cjerar)
    End If
    End If
   
If UserList(UserIndex).flags.CJerarquia = 1 Then
   Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "�" & "Ya sos cuarta jerarquia" & "�" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
End If
 
End Sub
 
Public Sub ExpulsarFaccionReal(ByVal UserIndex As Integer)
 
    UserList(UserIndex).flags.PJerarquia = 0
    UserList(UserIndex).flags.SJerarquia = 0
    UserList(UserIndex).flags.TJerarquia = 0
    UserList(UserIndex).flags.CJerarquia = 0
 
    UserList(UserIndex).Faccion.ArmadaReal = 0
    'Call PerderItemsFaccionarios(UserIndex)
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has sido expulsado de las tropas reales!!!." & FONTTYPE_FIGHT)
    'Desequipamos la armadura real si est� equipada
    If ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).Real = 1 Then Call Desequipar(UserIndex, UserList(UserIndex).Invent.ArmourEqpSlot)
End Sub
 
Public Sub ExpulsarFaccionCaos(ByVal UserIndex As Integer)
 
    UserList(UserIndex).flags.PJerarquia = 0
    UserList(UserIndex).flags.SJerarquia = 0
    UserList(UserIndex).flags.TJerarquia = 0
    UserList(UserIndex).flags.CJerarquia = 0
 
    UserList(UserIndex).Faccion.FuerzasCaos = 0
    'Call PerderItemsFaccionarios(UserIndex)
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has sido expulsado de la legi�n oscura!!!." & FONTTYPE_FIGHT)
    'Desequipamos la armadura real si est� equipada
    If ObjData(UserList(UserIndex).Invent.ArmourEqpObjIndex).Caos = 1 Then Call Desequipar(UserIndex, UserList(UserIndex).Invent.ArmourEqpSlot)
End Sub
 
Public Function TituloReal(ByVal UserIndex As Integer) As String
 
Select Case UserList(UserIndex).Faccion.RecompensasReal
    Case 0
        TituloReal = "Primera Jerarquia"
        UserList(UserIndex).flags.PJerarquia = 1
    Case 1
        TituloReal = "Segunda Jerarquia"
        UserList(UserIndex).flags.PJerarquia = 0
        UserList(UserIndex).flags.SJerarquia = 1
    Case 2
        TituloReal = "Tercera Jerarquia"
        UserList(UserIndex).flags.SJerarquia = 0
        UserList(UserIndex).flags.TJerarquia = 1
    Case Else
        TituloReal = "Cuarta Jerarquia"
        UserList(UserIndex).flags.TJerarquia = 0
        UserList(UserIndex).flags.CJerarquia = 1
End Select
 
End Function
 
Public Sub EnlistarCaos(ByVal UserIndex As Integer)
 
If Not Criminal(UserIndex) Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "�" & "Largate de aqui, bufon!!!!" & "�" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
    Exit Sub
End If
 
If UserList(UserIndex).Faccion.FuerzasCaos = 1 Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "�" & "Ya perteneces a la legi�n oscura!!!" & "�" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
    Exit Sub
End If
 
If UserList(UserIndex).Faccion.ArmadaReal = 1 Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "�" & "Las sombras reinaran en Argentum, largate de aqui estupido ciudadano.!!!" & "�" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
    Exit Sub
End If
 
'[Barrin 17-12-03] Si era miembro de la Armada Real no se puede enlistar
If UserList(UserIndex).Faccion.RecibioExpInicialReal = 1 Then 'Tomamos el valor de ah�: �Recibio la experiencia para entrar?
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "�" & "No permitir� que ning�n insecto real ingrese �Traidor del Rey!" & "�" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
    Exit Sub
End If
'[/Barrin]
 
If Not Criminal(UserIndex) Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "�" & "Ja ja ja tu no eres bienvenido aqui!!!" & "�" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
    Exit Sub
End If
 
If UserList(UserIndex).Faccion.CiudadanosMatados < 100 Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "�" & "Para unirte a nuestras fuerzas debes matar al menos 100 ciudadanos, solo has matado " & UserList(UserIndex).Faccion.CiudadanosMatados & "�" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
    Exit Sub
End If
 
If UserList(UserIndex).Stats.ELV < 25 Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "�" & "Para unirte a nuestras fuerzas debes ser al menos de nivel 25!!!" & "�" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
    Exit Sub
End If
 
 
If UserList(UserIndex).Faccion.Reenlistadas > 4 Then
    If UserList(UserIndex).Faccion.Reenlistadas = 200 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "�" & "Has sido expulsado de las fuerzas oscuras y durante tu rebeld�a has atacado a mi ej�rcito. Vete de aqu�!" & "�" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
    Else
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "�" & "Has sido expulsado de las fuerzas oscuras demasiadas veces!" & "�" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
    End If
    Exit Sub
End If
 
UserList(UserIndex).Faccion.Reenlistadas = UserList(UserIndex).Faccion.Reenlistadas + 1
UserList(UserIndex).Faccion.FuerzasCaos = 1
UserList(UserIndex).Faccion.RecompensasCaos = UserList(UserIndex).Faccion.CiudadanosMatados \ 100
 
Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "�" & "Bienvenido a al lado oscuro!!!, aqui tienes tu armadura. Por cada centena de ciudadanos que acabes te dar� un recompensa, buena suerte soldado!" & "�" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
 
If UserList(UserIndex).Faccion.RecibioArmaduraCaos = 0 Then
    Dim MiObj As Obj
    Dim MiObj2 As Obj
    MiObj.Amount = 20
    MiObj2.Amount = 10
   
    If UCase$(UserList(UserIndex).Raza) = "ENANO" Or UCase$(UserList(UserIndex).Raza) = "GNOMO" Then
        MiObj.ObjIndex = VestimentaLegionEnano
        Select Case UCase$(UserList(UserIndex).Clase)
            Case "MAGO"
                MiObj2.ObjIndex = TunicaEgregiaEnano
            Case Else
                MiObj2.ObjIndex = TunicaLobregaEnano
        End Select
    Else
        MiObj.ObjIndex = VestimentaLegionHumano
        Select Case UCase$(UserList(UserIndex).Clase)
            Case "MAGO"
                MiObj2.ObjIndex = TunicaEgregiaHumano
            Case "CLERIGO", "DRUIDA", "BARDO"
                MiObj2.ObjIndex = SacerdoteDemoniaco
            Case Else
                MiObj2.ObjIndex = TunicaLobregaHumano
        End Select
    End If
   
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If
    If Not MeterItemEnInventario(UserIndex, MiObj2) Then
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj2)
    End If
   
    UserList(UserIndex).Faccion.RecibioArmaduraCaos = 1
   
 
 
End If
 
If UserList(UserIndex).Faccion.RecibioExpInicialCaos = 0 Then
    UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + ExpAlUnirse
    If UserList(UserIndex).Stats.Exp > MAXEXP Then _
        UserList(UserIndex).Stats.Exp = MAXEXP
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has ganado " & ExpAlUnirse & " puntos de experiencia." & FONTTYPE_DAREXP)
    UserList(UserIndex).Faccion.RecibioExpInicialCaos = 1
    Call CheckUserLevel(UserIndex)
End If
 
 
Call LogEjercitoCaos(UserList(UserIndex).name)
 
End Sub
  
Public Sub RecompensaCaos(ByVal UserIndex As Integer)
 
Dim SjerarC As Obj
Dim TjerarC As Obj
Dim CjerarC As Obj
 
SjerarC.Amount = 1
SjerarC.ObjIndex = 865
 
TjerarC.Amount = 1
TjerarC.ObjIndex = 889
 
CjerarC.Amount = 1
CjerarC.ObjIndex = 866
 
 
 
If UserList(UserIndex).Faccion.CiudadanosMatados \ 100 = _
   UserList(UserIndex).Faccion.RecompensasCaos Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "�" & "Ya has recibido tu recompensa, mata 100 ciudadanos mas para recibir la proxima!!!" & "�" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
End If
 
If UserList(UserIndex).flags.PJerarquia = 1 Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "�" & "Aqui tienes tu recompensa noble guerrero!!!" & "�" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
    UserList(UserIndex).Faccion.RecompensasCaos = UserList(UserIndex).Faccion.RecompensasCaos + 1
    Call CheckUserLevel(UserIndex)
If Not MeterItemEnInventario(UserIndex, SjerarC) Then
    Call TirarItemAlPiso(UserList(UserIndex).Pos, SjerarC)
    End If
    End If
 
If UserList(UserIndex).flags.SJerarquia = 1 Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "�" & "Aqui tienes tu recompensa noble guerrero!!!" & "�" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
    UserList(UserIndex).Faccion.RecompensasCaos = UserList(UserIndex).Faccion.RecompensasCaos + 1
    Call CheckUserLevel(UserIndex)
If Not MeterItemEnInventario(UserIndex, TjerarC) Then
    Call TirarItemAlPiso(UserList(UserIndex).Pos, TjerarC)
    End If
    End If
 
If UserList(UserIndex).flags.TJerarquia = 1 Then
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "�" & "Aqui tienes tu recompensa noble guerrero!!!" & "�" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
    UserList(UserIndex).Faccion.RecompensasCaos = UserList(UserIndex).Faccion.RecompensasCaos + 1
    Call CheckUserLevel(UserIndex)
If Not MeterItemEnInventario(UserIndex, CjerarC) Then
    Call TirarItemAlPiso(UserList(UserIndex).Pos, CjerarC)
    End If
    End If
   
If UserList(UserIndex).flags.CJerarquia = 1 Then
   Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "�" & "Ya sos cuarta jerarquia" & "�" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
End If
 
 
End Sub
 
Public Function TituloCaos(ByVal UserIndex As Integer) As String
Select Case UserList(UserIndex).Faccion.RecompensasCaos
    Case 0
        TituloCaos = "Primera Jerarquia"
        UserList(UserIndex).flags.PJerarquia = 1
    Case 1
        TituloCaos = "Segunda Jerarquia"
        UserList(UserIndex).flags.PJerarquia = 0
        UserList(UserIndex).flags.SJerarquia = 1
    Case 2
        TituloCaos = "Tercera Jerarquia"
        UserList(UserIndex).flags.SJerarquia = 0
        UserList(UserIndex).flags.TJerarquia = 1
    Case Else
        TituloCaos = "Cuarta Jerarquia"
        UserList(UserIndex).flags.TJerarquia = 0
        UserList(UserIndex).flags.CJerarquia = 1
End Select
 
 
End Function
 
'[Barrin 17-12-03]
'Sub PerderItemsFaccionarios(ByVal UserIndex As Integer)
'Dim i As Byte
'Dim MiObj As Obj
'Dim ItemIndex As Integer
'
'For i = 1 To MAX_INVENTORY_SLOTS
'  ItemIndex = UserList(UserIndex).Invent.Object(i).ObjIndex
'  If ItemIndex > 0 Then
'         If ObjData(ItemIndex).Real = 1 Or ObjData(ItemIndex).Caos = 1 Then
'            Call QuitarUserInvItem(UserIndex, i, UserList(UserIndex).Invent.Object(i).Amount)
'            Call UpdateUserInv(False, UserIndex, i)
'            If ObjData(ItemIndex).ObjType = eOBJType.Armour Then
'                If ObjData(ItemIndex).Real = 1 Then UserList(UserIndex).Faccion.RecibioArmaduraReal = 0
'                If ObjData(ItemIndex).Caos = 1 Then UserList(UserIndex).Faccion.RecibioArmaduraCaos = 0
'            Else
'                UserList(UserIndex).Faccion.RecibioItemFaccionario = 0
'            End If
'         End If
'
'  End If
'Next i
'
'End Sub
'[/Barrin]

