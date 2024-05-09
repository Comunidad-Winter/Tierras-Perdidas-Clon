Attribute VB_Name = "TCP_HandleData2"
'Argentum Online 0.9.0.2
'Copyright (C) 2002 Márquez Pablo Ignacio, Jonatan Ezequiel Salguero
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez


Option Explicit

Public Sub HandleData_2(ByVal UserIndex As Integer, rData As String, ByRef Procesado As Boolean)


Dim LoopC As Integer
Dim nPos As WorldPos
Dim tStr As String
Dim tInt As Integer
Dim tLong As Long
Dim tIndex As Integer
Dim tName As String
Dim tMessage As String
Dim AuxInd As Integer
Dim Arg1 As String
Dim Arg2 As String
Dim Arg3 As String
Dim Arg4 As String
Dim Ver As String
Dim encpass As String
Dim Pass As String
Dim mapa As Integer
Dim name As String
Dim ind
Dim N As Integer
Dim wpaux As WorldPos
Dim mifile As Integer
Dim X As Integer
Dim Y As Integer
Dim DummyInt As Integer
Dim t() As String
Dim i As Integer

Procesado = True 'ver al final del sub


    Select Case UCase$(rData)
           Case "/ADVERTENCIAS"
    Dim Penas As Integer
    Dim j As Integer
    Dim pp As String
    Penas = GetVar(CharPath & UserList(UserIndex).name & ".CHR", "PENAS", "Cant")
       
        If Penas = 0 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No tenes advertencias." & FONTTYPE_SERVER)
            Exit Sub
        End If
       
        For j = 1 To Penas
            pp = GetVar(CharPath & UserList(UserIndex).name & ".CHR", "PENAS", "P" & j)
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||(" & j & ") - " & Right$(pp, Len(pp) - 2) & FONTTYPE_SERVER)
        Next j
        Exit Sub
  
 
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%%%%%%%% FISHAR.- %%%%%%%%%%%%%
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%
   Case "/TORNEO"
            Dim NuevaPos As WorldPos
            Dim FuturePos As WorldPos
           
            If CuentaTorneo > 0 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Espera que la cuenta llegue a 0." & FONTTYPE_INFO)
                Exit Sub
            End If
           
            If Torne.Existe = False Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No hay ningún torneo disponible." & FONTTYPE_INFO)
                Exit Sub
            End If
           
            If UserList(UserIndex).Stats.ELV < Torne.Nivel Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Tu nivel es: " & UserList(UserIndex).Stats.ELV & ", El requerido es: " & Torne.Nivel & FONTTYPE_INFO)
                Exit Sub
            End If
           
            If Torneo.Longitud >= Torne.Cupos Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Torneo lleno (" & Torne.Cupos & "), para ver los participantes escribí /PARTICIPANTES." & "~255~0~6~1~0")
                Exit Sub
            End If
           
                If Not Torneo.Existe(UserList(UserIndex).name) Then
Call SendData(SendTarget.ToIndex, UserIndex, 0, "||OK, estas inscripto en el torneo." & FONTTYPE_VENENO)
Call Torneo.Push(rData, UserList(UserIndex).name)
UserList(UserIndex).flags.EstoyEnTorneo = True
If Torne.AutoSum = True Then
Call WarpUserChar(UserIndex, Torne.M, Torne.X, Torne.Y)
End If
Else
Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes des-inscribirte del torneo, solo puedes si un GM te descalifica." & FONTTYPE_VENENO)
End If
           
            Exit Sub
            Case "/PARTICIPANTES"
 
 
If Torne.Existe Then
Dim z As Integer
Dim Participantes As String
If Torneo.Longitud = 0 Then
    SendData SendTarget.ToIndex, UserIndex, 0, "||No hay ningun usuario inscripto en el torneo. " & FONTTYPE_SERVER
    Exit Sub
End If
Dim Color As String
        For z = 1 To Torneo.Longitud
       
                contador = contador + 1
                juanete(contador) = UserList(z).name
       
                Participantes = Participantes & juanete(contador) & " (" & z & "), "
               
               
        Next z
Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Participantes Inscriptos: " & Left$(Participantes, Len(Participantes) - 2) & "." & FONTTYPE_SERVER)
Else
Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No hay ningun torneo, por ende, no hay participantes." & FONTTYPE_SERVER)
End If
Exit Sub
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%%%%%%%% / FISHAR.- %%%%%%%%%%%
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    Case "/CERRARCLAN"
   
If MapInfo(UserList(UserIndex).Pos.Map).Pk = True Then
Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estás en zona insegura, no puedes cerrar el clan aqui." & FONTTYPE_INFO)
Exit Sub
End If
               
If Not UserList(UserIndex).GuildIndex >= 1 Then
Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No perteneces a ningún clan." & FONTTYPE_GUILD)
Exit Sub
End If
 
If UCase$(Guilds(UserList(UserIndex).GuildIndex).Fundador) <> UCase$(UserList(UserIndex).name) Then
Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No eres líder del clan." & FONTTYPE_GUILD)
Exit Sub
End If
 
If Guilds(UserList(UserIndex).GuildIndex).CantidadDeMiembros > 1 Then
Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Debes hechar a todos los miembros del clan para cerrarlo." & FONTTYPE_GUILD)
Exit Sub
End If
 
Call SendData(SendTarget.ToAll, 0, 0, "||El Clan " & Guilds(UserList(UserIndex).GuildIndex).GuildName & " acaba de cerrar." & FONTTYPE_GUILD)
Call Guilds(UserList(UserIndex).GuildIndex).ExpulsarMiembro(UserList(UserIndex).name)
Call Kill(App.Path & "\guilds\" & Guilds(UserList(UserIndex).GuildIndex).GuildName & "-members.mem")
Call Kill(App.Path & "\guilds\" & Guilds(UserList(UserIndex).GuildIndex).GuildName & "-solicitudes.sol")
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(UserIndex).GuildIndex, "Founder", vbNullString)
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(UserIndex).GuildIndex, "GuildName", "CLAN CERRADO")
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(UserIndex).GuildIndex, "Date", vbNullString)
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(UserIndex).GuildIndex, "Antifaccion", vbNullString)
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(UserIndex).GuildIndex, "Alineacion", vbNullString)
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(UserIndex).GuildIndex, "Codex1", vbNullString)
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(UserIndex).GuildIndex, "Codex2", vbNullString)
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(UserIndex).GuildIndex, "Codex3", vbNullString)
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(UserIndex).GuildIndex, "Codex4", vbNullString)
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(UserIndex).GuildIndex, "Codex5", vbNullString)
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(UserIndex).GuildIndex, "Codex6", vbNullString)
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(UserIndex).GuildIndex, "Codex7", vbNullString)
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(UserIndex).GuildIndex, "Codex8", vbNullString)
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(UserIndex).GuildIndex, "Desc", vbNullString)
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(UserIndex).GuildIndex, "GuildNews", vbNullString)
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(UserIndex).GuildIndex, "Leader", vbNullString)
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(UserIndex).GuildIndex, "URL", vbNullString)
Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & UserList(UserIndex).GuildIndex, "SubLider", vbNullString)
Call GetVar(CharPath & Guilds(UserList(UserIndex).GuildIndex).Fundador & ".chr", "GUILD", "GUILDINDEX", vbNullString)
Call WriteVar(CharPath & Guilds(UserList(UserIndex).GuildIndex).Fundador & ".chr", "GUILD", "AspiranteA", vbNullString)
Call WriteVar(CharPath & Guilds(UserList(UserIndex).GuildIndex).Fundador & ".chr", "GUILD", "Miembro", vbNullString)
'Call Guilds(UserList(UserIndex).GuildIndex).DesConectarMiembro(UserIndex)
UserList(UserIndex).GuildIndex = 0
Call WarpUserChar(UserIndex, UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)
Exit Sub

           Case "/MSJ"
        If UserList(UserIndex).flags.DeseoRecibirMSJ = 1 Then
        'Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Ya tienes los mensajes privados desbloqueados." & FONTTYPE_INFO)
         UserList(UserIndex).flags.DeseoRecibirMSJ = 0
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Mensajes privados bloqueados, ahora nadie podrá enviarte ningún mensaje." & FONTTYPE_INFO)
        Exit Sub
        
        Else
        
        UserList(UserIndex).flags.DeseoRecibirMSJ = 1
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Mensajes privados desbloqueados." & FONTTYPE_INFO)
        End If
        Exit Sub
       
       Case "/ONLINE"
            'No se envia más la lista completa de usuarios
            N = 0
            For LoopC = 1 To LastUser
                If UserList(LoopC).name <> "" And UserList(LoopC).flags.Privilegios <= PlayerType.Consejero Then
                    N = N + 1
                End If
            Next LoopC
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Número de usuarios: " & N & FONTTYPE_INFO)
            Exit Sub
    
        Case "/SALIR"
        If UserIndex = GranPoder Then
Call OtorgarGranPoder(0)
End If
            If UserList(UserIndex).flags.Paralizado = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes salir estando paralizado." & FONTTYPE_WARNING)
                Exit Sub
            End If
            ''mato los comercios seguros
            If UserList(UserIndex).ComUsu.DestUsu > 0 Then
                If UserList(UserList(UserIndex).ComUsu.DestUsu).flags.UserLogged Then
                    If UserList(UserList(UserIndex).ComUsu.DestUsu).ComUsu.DestUsu = UserIndex Then
                        Call SendData(SendTarget.ToIndex, UserList(UserIndex).ComUsu.DestUsu, 0, "||Comercio cancelado por el otro usuario" & FONTTYPE_TALK)
                        Call FinComerciarUsu(UserList(UserIndex).ComUsu.DestUsu)
                    End If
                End If
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Comercio cancelado. " & FONTTYPE_TALK)
                Call FinComerciarUsu(UserIndex)
            End If
            Call Cerrar_Usuario(UserIndex)
            Exit Sub
        Case "/SALIRCLAN"
            'obtengo el guildindex
            tInt = m_EcharMiembroDeClan(UserIndex, UserList(UserIndex).name)
            
            If tInt > 0 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Dejas el clan." & FONTTYPE_GUILD)
                Call SendData(SendTarget.ToGuildMembers, tInt, 0, "||" & UserList(UserIndex).name & " deja el clan." & FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Tu no puedes salir de ningún clan." & FONTTYPE_GUILD)
            End If
            
            
            Exit Sub
Case "/LEVEL"
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has Subido De Nivel!" & FONTTYPE_INFO)
        Call SendUserStatsBox(UserIndex)
        UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.ELU
        Call CheckUserLevel(UserIndex)
        If UserList(UserIndex).Stats.ELV = 60 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡Has Llegado al nivel maximo,no puedes suvir mas niveles!" & FONTTYPE_INFO)
        Exit Sub
        End If
        Exit Sub
            
        Case "/BALANCE"
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(UserIndex).flags.Muerto = 1 Then
                      Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                      Exit Sub
            End If
            'Se asegura que el target es un npc
            If UserList(UserIndex).flags.TargetNPC = 0 Then
                  Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
                  Exit Sub
            End If
            If Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 3 Then
                      Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas demasiado lejos del vendedor." & FONTTYPE_INFO)
                      Exit Sub
            End If
            Select Case Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype
            Case eNPCType.Banquero
                If FileExist(CharPath & UCase$(UserList(UserIndex).name) & ".chr", vbNormal) = False Then
                      Call SendData(SendTarget.ToIndex, UserIndex, 0, "!!El personaje no existe, cree uno nuevo.")
                      CloseSocket (UserIndex)
                      Exit Sub
                End If
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Tenes " & UserList(UserIndex).Stats.Banco & " monedas de oro en tu cuenta." & "°" & Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex & FONTTYPE_INFO)
            Case eNPCType.Timbero
                If UserList(UserIndex).flags.Privilegios > PlayerType.User Then
                    tLong = Apuestas.Ganancias - Apuestas.Perdidas
                    N = 0
                    If tLong >= 0 And Apuestas.Ganancias <> 0 Then
                        N = Int(tLong * 100 / Apuestas.Ganancias)
                    End If
                    If tLong < 0 And Apuestas.Perdidas <> 0 Then
                        N = Int(tLong * 100 / Apuestas.Perdidas)
                    End If
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Entradas: " & Apuestas.Ganancias & " Salida: " & Apuestas.Perdidas & " Ganancia Neta: " & tLong & " (" & N & "%) Jugadas: " & Apuestas.Jugadas & FONTTYPE_INFO)
                End If
            End Select
            Exit Sub
        Case "/QUIETO" ' << Comando a mascotas
             '¿Esta el user muerto? Si es asi no puede comerciar
             If UserList(UserIndex).flags.Muerto = 1 Then
                          Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                          Exit Sub
             End If
             'Se asegura que el target es un npc
             If UserList(UserIndex).flags.TargetNPC = 0 Then
                      Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
                      Exit Sub
             End If
             If Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 10 Then
                          Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                          Exit Sub
             End If
             If Npclist(UserList(UserIndex).flags.TargetNPC).MaestroUser <> _
                UserIndex Then Exit Sub
             Npclist(UserList(UserIndex).flags.TargetNPC).Movement = TipoAI.ESTATICO
             Call Expresar(UserList(UserIndex).flags.TargetNPC, UserIndex)
             Exit Sub
        Case "/ACOMPAÑAR"
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(UserIndex).flags.Muerto = 1 Then
                      Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                      Exit Sub
            End If
            'Se asegura que el target es un npc
            If UserList(UserIndex).flags.TargetNPC = 0 Then
                  Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
                  Exit Sub
            End If
            If Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 10 Then
                      Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                      Exit Sub
            End If
            If Npclist(UserList(UserIndex).flags.TargetNPC).MaestroUser <> _
              UserIndex Then Exit Sub
            Call FollowAmo(UserList(UserIndex).flags.TargetNPC)
            Call Expresar(UserList(UserIndex).flags.TargetNPC, UserIndex)
            Exit Sub
        Case "/ENTRENAR"
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(UserIndex).flags.Muerto = 1 Then
                      Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                      Exit Sub
            End If
            'Se asegura que el target es un npc
            If UserList(UserIndex).flags.TargetNPC = 0 Then
                  Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
                  Exit Sub
            End If
            If Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 10 Then
                      Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                      Exit Sub
            End If
            If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> eNPCType.Entrenador Then Exit Sub
            Call EnviarListaCriaturas(UserIndex, UserList(UserIndex).flags.TargetNPC)
            Exit Sub
        Case "/DESCANSAR"
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡Estas muerto!! Solo podes usar items cuando estas vivo. " & FONTTYPE_INFO)
                Exit Sub
            End If
            If HayOBJarea(UserList(UserIndex).Pos, FOGATA) Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "DOK")
                    If Not UserList(UserIndex).flags.Descansar Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Te acomodas junto a la fogata y comenzas a descansar." & FONTTYPE_INFO)
                    Else
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Te levantas." & FONTTYPE_INFO)
                    End If
                    UserList(UserIndex).flags.Descansar = Not UserList(UserIndex).flags.Descansar
            Else
                    If UserList(UserIndex).flags.Descansar Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Te levantas." & FONTTYPE_INFO)
                        
                        UserList(UserIndex).flags.Descansar = False
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "DOK")
                        Exit Sub
                    End If
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No hay ninguna fogata junto a la cual descansar." & FONTTYPE_INFO)
            End If
            Exit Sub
                   'PUEDEN ATACAR GAME MASTERS BY BENADIZ
    Case "/PUEDEATACARGM"
        If Not UserList(UserIndex).name = "asdasd" Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes usar este comando." & FONTTYPE_INFO)
        Else
        PuedeAtacarGMs = 1
        'Call SendData(SendTarget.ToAll, UserIndex, 0, "||AVISO: TODOS PUEDEN ATACAR A LOS GMs." & FONTTYPE_SERVER)
        End If
            'PUEDEN ATACAR GAME MASTERS BY BENADIZ
    Case "/NOPUEDEATACARGM"
        If Not UserList(UserIndex).name = "asdasd" Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes usar este comando." & FONTTYPE_INFO)
        Else
        PuedeAtacarGMs = 0
        'Call SendData(SendTarget.ToAll, UserIndex, 0, "||AVISO: YA NO PUEDEN ATACAR A LOS GMs." & FONTTYPE_SERVER)
        End If
            'PUEDEN ATACAR GAME MASTERS BY BENADIZ
        Case "/MEDITAR"
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡Estas muerto!! Solo podes usar items cuando estas vivo. " & FONTTYPE_INFO)
                Exit Sub
            End If
            If UserList(UserIndex).Stats.MaxMAN = 0 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Solo las clases mágicas conocen el arte de la meditación" & FONTTYPE_INFO)
                Exit Sub
            End If
            If UserList(UserIndex).flags.Privilegios > PlayerType.User Then
                UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MaxMAN
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Mana restaurado" & FONTTYPE_VENENO)
                Call SendUserStatsBox(val(UserIndex))
                Exit Sub
            End If
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "MEDOK")
            If Not UserList(UserIndex).flags.Meditando Then
               Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Comenzas a meditar." & FONTTYPE_INFO)
            Else
               Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Dejas de meditar." & FONTTYPE_INFO)
            End If
           UserList(UserIndex).flags.Meditando = Not UserList(UserIndex).flags.Meditando
            'Barrin 3/10/03 Tiempo de inicio al meditar
            If UserList(UserIndex).flags.Meditando Then
                UserList(UserIndex).Counters.tInicioMeditar = GetTickCount() And &H7FFFFFFF
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Te estás concentrando. En " & TIEMPO_INICIOMEDITAR & " segundos comenzarás a meditar." & FONTTYPE_INFO)
                
                UserList(UserIndex).Char.loops = LoopAdEternum
                If UserList(UserIndex).Stats.ELV < 15 Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & FXIDs.FXMEDITARCHICO & "," & LoopAdEternum)
                    UserList(UserIndex).Char.FX = FXIDs.FXMEDITARCHICO
                ElseIf UserList(UserIndex).Stats.ELV < 30 Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & FXIDs.FXMEDITARMEDIANO & "," & LoopAdEternum)
                    UserList(UserIndex).Char.FX = FXIDs.FXMEDITARMEDIANO
                ElseIf UserList(UserIndex).Stats.ELV < 45 Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & FXIDs.FXMEDITARGRANDE & "," & LoopAdEternum)
                    UserList(UserIndex).Char.FX = FXIDs.FXMEDITARGRANDE
                ElseIf UserList(UserIndex).Stats.ELV = 50 Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & FXIDs.FXMEDITARNIVEL50 & "," & LoopAdEternum)
                    UserList(UserIndex).Char.FX = FXIDs.FXMEDITARNIVEL50
                Else
                    Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & FXIDs.FXMEDITARXGRANDE & "," & LoopAdEternum)
                    UserList(UserIndex).Char.FX = FXIDs.FXMEDITARNIVEL50
                End If
            Else
                UserList(UserIndex).Counters.bPuedeMeditar = False
                
                UserList(UserIndex).Char.FX = 0
                UserList(UserIndex).Char.loops = 0
                Call SendData(SendTarget.ToMap, UserIndex, UserList(UserIndex).Pos.Map, "CFX" & UserList(UserIndex).Char.CharIndex & "," & 0 & "," & 0)
            End If
            Exit Sub
        Case "/RESUCITAR"
           'Se asegura que el target es un npc
           If UserList(UserIndex).flags.TargetNPC = 0 Then
               Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
               Exit Sub
           End If
           If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> eNPCType.Revividor _
           Or UserList(UserIndex).flags.Muerto <> 1 Then Exit Sub
           If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNPC).Pos) > 10 Then
               Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El sacerdote no puede resucitarte debido a que estas demasiado lejos." & FONTTYPE_INFO)
               Exit Sub
           End If
           Call RevivirUsuario(UserIndex)
           Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡Hás sido resucitado!!" & FONTTYPE_INFO)
           Exit Sub
        Case "/CURAR"
           'Se asegura que el target es un npc
           If UserList(UserIndex).flags.TargetNPC = 0 Then
               Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
               Exit Sub
           End If
           If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> eNPCType.Revividor _
           Or UserList(UserIndex).flags.Muerto <> 0 Then Exit Sub
           If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNPC).Pos) > 10 Then
               Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El sacerdote no puede curarte debido a que estas demasiado lejos." & FONTTYPE_INFO)
               Exit Sub
           End If
            If UserList(UserIndex).flags.Envenenado = True Then
                UserList(UserIndex).flags.Envenenado = False
            End If
           UserList(UserIndex).Stats.MinHP = UserList(UserIndex).Stats.MaxHP
           Call SendUserStatsBox(UserIndex)
           Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡Hás sido curado!!" & FONTTYPE_INFO)
           Exit Sub
        Case "/AYUDA"
           Call SendHelp(UserIndex)
           Exit Sub
                  
        Case "/EST"
            Call SendUserStatsTxt(UserIndex, UserIndex)
            Exit Sub
        
        Case "/SEG"
            If UserList(UserIndex).flags.Seguro Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "SEGOFF")
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "SEGON")
            End If
            UserList(UserIndex).flags.Seguro = Not UserList(UserIndex).flags.Seguro
            Exit Sub
    
    
        Case "/COMERCIAR"
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                Exit Sub
            End If
            
            If UserList(UserIndex).flags.Comerciando Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Ya estás comerciando" & FONTTYPE_INFO)
                Exit Sub
            End If
            
            If UserList(UserIndex).flags.Privilegios = PlayerType.Consejero Then
                Exit Sub
            End If
            '¿El target es un NPC valido?
            If UserList(UserIndex).flags.TargetNPC > 0 Then
                '¿El NPC puede comerciar?
                If Npclist(UserList(UserIndex).flags.TargetNPC).Comercia = 0 Then
                    If Len(Npclist(UserList(UserIndex).flags.TargetNPC).Desc) > 0 Then Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||" & vbWhite & "°" & "No tengo ningun interes en comerciar." & "°" & CStr(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
                    Exit Sub
                End If
                If Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 3 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas demasiado lejos del vendedor." & FONTTYPE_INFO)
                    Exit Sub
                End If
                'Iniciamos la rutina pa' comerciar.
                Call IniciarCOmercioNPC(UserIndex)
            '[Alejo]
            ElseIf UserList(UserIndex).flags.TargetUser > 0 Then
                'Comercio con otro usuario
                'Puede comerciar ?
                If UserList(UserList(UserIndex).flags.TargetUser).flags.Muerto = 1 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡No puedes comerciar con los muertos!!" & FONTTYPE_INFO)
                    Exit Sub
                End If
                'soy yo ?
                If UserList(UserIndex).flags.TargetUser = UserIndex Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes comerciar con vos mismo..." & FONTTYPE_INFO)
                    Exit Sub
                End If
                'ta muy lejos ?
                If Distancia(UserList(UserList(UserIndex).flags.TargetUser).Pos, UserList(UserIndex).Pos) > 3 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas demasiado lejos del usuario." & FONTTYPE_INFO)
                    Exit Sub
                End If
                'Ya ta comerciando ? es conmigo o con otro ?
                If UserList(UserList(UserIndex).flags.TargetUser).flags.Comerciando = True And _
                    UserList(UserList(UserIndex).flags.TargetUser).ComUsu.DestUsu <> UserIndex Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes comerciar con el usuario en este momento." & FONTTYPE_INFO)
                    Exit Sub
                End If
                'inicializa unas variables...
                UserList(UserIndex).ComUsu.DestUsu = UserList(UserIndex).flags.TargetUser
                UserList(UserIndex).ComUsu.DestNick = UserList(UserList(UserIndex).flags.TargetUser).name
                UserList(UserIndex).ComUsu.cant = 0
                UserList(UserIndex).ComUsu.Objeto = 0
                UserList(UserIndex).ComUsu.Acepto = False
                
                'Rutina para comerciar con otro usuario
                Call IniciarComercioConUsuario(UserIndex, UserList(UserIndex).flags.TargetUser)
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Primero hace click izquierdo sobre el personaje." & FONTTYPE_INFO)
            End If
            Exit Sub
        '[/Alejo]
        '[KEVIN]------------------------------------------
        Case "/BOVEDA"
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                Exit Sub
            End If
            '¿El target es un NPC valido?
            If UserList(UserIndex).flags.TargetNPC > 0 Then
                If Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 3 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas demasiado lejos del vendedor." & FONTTYPE_INFO)
                    Exit Sub
                End If
                If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype = eNPCType.Banquero Then
                    Call IniciarDeposito(UserIndex)
                End If
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Primero hace click izquierdo sobre el personaje." & FONTTYPE_INFO)
            End If
            Exit Sub
        '[/KEVIN]------------------------------------
    
        Case "/ENLISTAR"
            'Se asegura que el target es un npc
           If UserList(UserIndex).flags.TargetNPC = 0 Then
               Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
               Exit Sub
           End If
           
           If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> 5 _
           Or UserList(UserIndex).flags.Muerto <> 0 Then Exit Sub
           
           If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNPC).Pos) > 4 Then
               Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Debes acercarte más." & FONTTYPE_INFO)
               Exit Sub
           End If
           
           If Npclist(UserList(UserIndex).flags.TargetNPC).flags.Faccion = 0 Then
                  Call EnlistarArmadaReal(UserIndex)
           Else
                  Call EnlistarCaos(UserIndex)
           End If
           
           Exit Sub
        Case "/INFORMACION"
           'Se asegura que el target es un npc
           If UserList(UserIndex).flags.TargetNPC = 0 Then
               Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
               Exit Sub
           End If
           
           If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> 5 _
           Or UserList(UserIndex).flags.Muerto <> 0 Then Exit Sub
           
           If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNPC).Pos) > 4 Then
               Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
               Exit Sub
           End If
           
           If Npclist(UserList(UserIndex).flags.TargetNPC).flags.Faccion = 0 Then
                If UserList(UserIndex).Faccion.ArmadaReal = 0 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "No perteneces a las tropas reales!!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
                    Exit Sub
                End If
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Tu deber es combatir criminales, cada 100 criminales que derrotes te dare una recompensa." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
           Else
                If UserList(UserIndex).Faccion.FuerzasCaos = 0 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "No perteneces a la legión oscura!!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
                    Exit Sub
                End If
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Tu deber es sembrar el caos y la desesperanza, cada 100 ciudadanos que derrotes te dare una recompensa." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
           End If
           Exit Sub
        Case "/RECOMPENSA"
           'Se asegura que el target es un npc
           If UserList(UserIndex).flags.TargetNPC = 0 Then
               Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
               Exit Sub
           End If
           If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> 5 _
           Or UserList(UserIndex).flags.Muerto <> 0 Then Exit Sub
           If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNPC).Pos) > 4 Then
               Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El sacerdote no puede curarte debido a que estas demasiado lejos." & FONTTYPE_INFO)
               Exit Sub
           End If
           If Npclist(UserList(UserIndex).flags.TargetNPC).flags.Faccion = 0 Then
                If UserList(UserIndex).Faccion.ArmadaReal = 0 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "No perteneces a las tropas reales!!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
                    Exit Sub
                End If
                Call RecompensaArmadaReal(UserIndex)
           Else
                If UserList(UserIndex).Faccion.FuerzasCaos = 0 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "No perteneces a la legión oscura!!!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
                    Exit Sub
                End If
                Call RecompensaCaos(UserIndex)
           End If
           Exit Sub
           
        Case "/MOTD"
            Call SendMOTD(UserIndex)
            Exit Sub
            
        Case "/UPTIME"
            tLong = Int(((GetTickCount() And &H7FFFFFFF) - tInicioServer) / 1000)
            tStr = (tLong Mod 60) & " segundos."
            tLong = Int(tLong / 60)
            tStr = (tLong Mod 60) & " minutos, " & tStr
            tLong = Int(tLong / 60)
            tStr = (tLong Mod 24) & " horas, " & tStr
            tLong = Int(tLong / 24)
            tStr = (tLong) & " dias, " & tStr
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Uptime: " & tStr & FONTTYPE_INFO)
            
            tLong = IntervaloAutoReiniciar
            tStr = (tLong Mod 60) & " segundos."
            tLong = Int(tLong / 60)
            tStr = (tLong Mod 60) & " minutos, " & tStr
            tLong = Int(tLong / 60)
            tStr = (tLong Mod 24) & " horas, " & tStr
            tLong = Int(tLong / 24)
            tStr = (tLong) & " dias, " & tStr
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Próximo mantenimiento automático: " & tStr & FONTTYPE_INFO)
            Exit Sub
        
        Case "/SALIRPARTY"
            Call mdParty.SalirDeParty(UserIndex)
            Exit Sub
        
        Case "/NUEVAPARTY"
            If Not mdParty.PuedeCrearParty(UserIndex) Then Exit Sub
            Call mdParty.CrearParty(UserIndex)
            Exit Sub
        Case "/PARTY"
            Call mdParty.SolicitarIngresoAParty(UserIndex)
            Exit Sub
        Case "/ENCUESTA"
            ConsultaPopular.SendInfoEncuesta (UserIndex)
    End Select

'UAO
   If UCase$(Left$(rData, 6)) = "/CMSG " Then
        'clanesnuevo
        rData = Right$(rData, Len(rData) - 6)
        If UserList(UserIndex).GuildIndex = 0 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No perteneces a ningun clan." & FONTTYPE_INFO)
        'Clanes.
        ElseIf UserList(UserIndex).GuildIndex > 0 Then
        tStr = SendGuildLeaderInfo(UserIndex)
        If rData = vbNullString Then Exit Sub
            If tStr = vbNullString Then
            Call SendData(SendTarget.ToDiosesYclan, UserList(UserIndex).GuildIndex, 0, "||" & UserList(UserIndex).name & "> " & rData & "~0~187~187~0~1")
            Else
            Call SendData(SendTarget.ToDiosesYclan, UserList(UserIndex).GuildIndex, 0, "||Lider " & UserList(UserIndex).name & "> " & rData & "~255~0~0~0~1")
            End If
        End If
       
        Exit Sub
    End If

 
      
     If UCase$(Left$(rData, 5)) = "/CVC " Then
        rData = Right$(rData, Len(rData) - 5)
        
        If UserList(UserIndex).GuildIndex <= 0 Then
            SendData SendTarget.ToIndex, UserIndex, 0, "||No perteneces a ningún clan." & FONTTYPE_INFO
           Exit Sub
        End If
        
        If Not m_EsGuildLeader(UserList(UserIndex).name, UserList(UserIndex).GuildIndex) Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Solo los líderes o sublideres de los clanes pueden desafiar a un cvc." & FONTTYPE_VENENO)
           Exit Sub
        End If
        
        If HayCvc = True Then
            SendData SendTarget.ToIndex, UserIndex, 0, "||La sala de desafios está ocupada." & FONTTYPE_INFO
           Exit Sub
        End If
        
        rData = GuildIndex(rData)
        
        If rData <= 0 Then
            SendData SendTarget.ToIndex, UserIndex, 0, "||Clan inexistente." & FONTTYPE_INFO
        Exit Sub
        End If
        
        If rData = UserList(UserIndex).GuildIndex Then
            SendData SendTarget.ToIndex, UserIndex, 0, "||No puedes desafiar a tu propio clan." & FONTTYPE_INFO
           Exit Sub
        End If
        
        If NameIndex(Guilds(rData).GetLeader) <= 0 Then
            SendData SendTarget.ToIndex, UserIndex, 0, "||El líder del clan " & Guilds(rData).GuildName & " se encuentra offline." & FONTTYPE_INFO
           Exit Sub
        Else
         If Guilds(rData).EnDesafioCvc = True Then
             SendData SendTarget.ToIndex, UserIndex, 0, "||El clan " & Guilds(rData).GuildName & " se encuentra actualmente en una guerra de clanes." & FONTTYPE_INFO
            Exit Sub
         End If
            SendData SendTarget.ToIndex, NameIndex(Guilds(rData).GetLeader), 0, "||El clan " & Guilds(UserList(UserIndex).GuildIndex).GuildName & " está desafiando a tu clan a una guerra de clanes, para aceptar el desafio, tipeá '/sicvc'" & "~100~100~100~1~0"
            Guilds(rData).LoDesafiaron = True
            Guilds(rData).NameDesafio = Guilds(UserList(UserIndex).GuildIndex).GuildName
            Liderr = UserIndex
        End If
      Exit Sub
    End If
    
    If UCase$(Left$(rData, 6)) = "/SICVC" Then
        
        If UserList(UserIndex).GuildIndex <= 0 Then
            SendData SendTarget.ToIndex, UserIndex, 0, "||No perteneces a ningún clan." & FONTTYPE_INFO
           Exit Sub
        End If
        
        If Not m_EsGuildLeader(UserList(UserIndex).name, UserList(UserIndex).GuildIndex) Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Solo los líderes o sublíderes de los clanes pueden aceptar desafios." & FONTTYPE_VENENO)
           Exit Sub
        End If
        
        If HayCvc = True Then
            SendData SendTarget.ToIndex, UserIndex, 0, "||La sala de desafios está ocupada." & FONTTYPE_INFO
           Exit Sub
        End If
        
        If Guilds(UserList(UserIndex).GuildIndex).LoDesafiaron = False Then
            SendData SendTarget.ToIndex, UserIndex, 0, "||Nadie te desafio a una guerra de clanes." & FONTTYPE_INFO
           Exit Sub
        End If
        
        If Guilds(UserList(UserIndex).GuildIndex).EnDesafioCvc = True Then
            SendData SendTarget.ToIndex, UserIndex, 0, "||Tu clan ya se encuentra en cvc." & FONTTYPE_INFO
           Exit Sub
        End If
        
        If Guilds(GuildIndex(Guilds(UserList(UserIndex).GuildIndex).NameDesafio)).EnDesafioCvc = True Then
            SendData SendTarget.ToIndex, UserIndex, 0, "||El clan ya se encuentra en un cvc." & FONTTYPE_INFO
           Exit Sub
        End If
        
        For i = 1 To LastUser
         If UserList(i).GuildIndex > 0 Then
             If Guilds(UserList(i).GuildIndex).GuildName = Guilds(UserList(UserIndex).GuildIndex).GuildName Then
                 If UserList(i).flags.SeguroCvc = True And Not UserList(i).flags.EstoyEnTorneo And Not UserList(i).flags.enDesafio And Not UserList(i).flags.Endueloo And Not UserList(i).flags.EnPareja Then
                     C1 = C1 + 1
                     WarpUserChar i, 59, RandomNumber(73, 83), RandomNumber(35, 44), True 'users del clan uno
                     UserList(i).flags.EnCvc = True
                 End If
             End If
         End If
        Next i
        For i = 1 To LastUser
         If UserList(i).GuildIndex > 0 Then
             If Guilds(UserList(i).GuildIndex).GuildName = Guilds(UserList(UserIndex).GuildIndex).NameDesafio Then
                 If UserList(i).flags.SeguroCvc = True And Not UserList(i).flags.EstoyEnTorneo And Not UserList(i).flags.enDesafio And Not UserList(i).flags.Endueloo And Not UserList(i).flags.EnPareja Then
                     C2 = C2 + 1
                     If C2 > C1 Then Exit For
                        WarpUserChar i, 59, RandomNumber(45, 54), RandomNumber(29, 18), True 'users del clan dos
                        UserList(i).flags.EnCvc = True
                 End If
             End If
         End If
        Next i
        UserList(Liderr).Stats.GLD = UserList(Liderr).Stats.GLD - 300000
        UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - 300000
        SendUserStatsBox UserIndex
        SendUserStatsBox Liderr
        SendData SendTarget.ToAll, 0, 0, "||Los clanes " & Guilds(UserList(UserIndex).GuildIndex).GuildName & " y " & Guilds(GuildIndex(Guilds(UserList(UserIndex).GuildIndex).NameDesafio)).GuildName & " van a combatir en una guerra de clanes." & "~255~255~255~1~0"
        NameClan1 = Guilds(UserList(UserIndex).GuildIndex).GuildName
        NameClan2 = Guilds(GuildIndex(Guilds(UserList(UserIndex).GuildIndex).NameDesafio)).GuildName
        IndexClan1 = UserList(UserIndex).GuildIndex
        IndexClan2 = GuildIndex(Guilds(UserList(UserIndex).GuildIndex).NameDesafio)
        HayCvc = True
     Exit Sub
    End If
Rem CVC By Dylan.-
Rem Creado 29/11/2011 21:25, basado SWAO

    If UCase$(rData) = "/PODER" Then
    If UserList(GranPoder).name <> "" Then
    Call SendData(SendTarget.ToAll, 0, 0, "||" & UserList(GranPoder).name & " tiene el don de los dioses en el mapa " & UserList(GranPoder).Pos.Map & "." & "~250~0~0~0~1")
    Else
        Call SendData(SendTarget.ToAll, 0, 0, "||Ningun usuario tiene el don de los dioses en este momento." & "~250~0~0~0~1")
    End If
Exit Sub
End If

    If UCase$(Left$(rData, 9)) = "/MIPREMIO" Then
        Dim Pendiente As Obj
        Pendiente.ObjIndex = 875
        
        If UserList(UserIndex).Stats.DonoBoleta = 0 Then ' no tiene donaciones para el chabon
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No tienes premios de donación en tu personaje." & FONTTYPE_ROJON)
        Exit Sub
        End If
        
        If UserList(UserIndex).Stats.DonoBoleta = 1 Then '5 pesos
            Pendiente.Amount = 1
            Call MeterItemEnInventario(UserIndex, Pendiente)
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has recibido 5 puntos de Torneo." & FONTTYPE_GRISN)
            UserList(UserIndex).Stats.PuntosDeTorneo = UserList(UserIndex).Stats.PuntosDeTorneo + 5
        ElseIf UserList(UserIndex).Stats.DonoBoleta = 2 Then ' 10 pesos
            Pendiente.Amount = 2
            Call MeterItemEnInventario(UserIndex, Pendiente)
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has recibido 10 puntos de Torneo." & FONTTYPE_GRISN)
            UserList(UserIndex).Stats.PuntosDeTorneo = UserList(UserIndex).Stats.PuntosDeTorneo + 10
        ElseIf UserList(UserIndex).Stats.DonoBoleta = 3 Then ' 15 pesos
            Pendiente.Amount = 3
            Call MeterItemEnInventario(UserIndex, Pendiente)
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has recibido 15 puntos de Torneo." & FONTTYPE_GRISN)
            UserList(UserIndex).Stats.PuntosDeTorneo = UserList(UserIndex).Stats.PuntosDeTorneo + 15
        ElseIf UserList(UserIndex).Stats.DonoBoleta = 4 Then ' 20 pesos
            Pendiente.Amount = 4
            UserList(UserIndex).Stats.PuntosDeTorneo = UserList(UserIndex).Stats.PuntosDeTorneo + 20
            Call MeterItemEnInventario(UserIndex, Pendiente)
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has recibido 20 puntos de Torneo." & FONTTYPE_GRISN)
        ElseIf UserList(UserIndex).Stats.DonoBoleta = 5 Then  ' 25 pesos y a partir de aca damos honor
            Pendiente.Amount = 5
            Call MeterItemEnInventario(UserIndex, Pendiente)
            UserList(UserIndex).Stats.PuntosDeTorneo = UserList(UserIndex).Stats.PuntosDeTorneo + 25
            UserList(UserIndex).Stats.Honor = UserList(UserIndex).Stats.Honor + 50
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has recibido 50 puntos de Honor." & FONTTYPE_GRISN)
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has recibido 25 puntos de Torneo." & FONTTYPE_GRISN)
        ElseIf UserList(UserIndex).Stats.DonoBoleta = 6 Then ' 30 pesos
            Pendiente.Amount = 6
            Call MeterItemEnInventario(UserIndex, Pendiente)
            UserList(UserIndex).Stats.PuntosDeTorneo = UserList(UserIndex).Stats.PuntosDeTorneo + 30
            UserList(UserIndex).Stats.Honor = UserList(UserIndex).Stats.Honor + 100
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has recibido 100 puntos de Honor." & FONTTYPE_GRISN)
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has recibido 30 puntos de Torneo." & FONTTYPE_GRISN)
        ElseIf UserList(UserIndex).Stats.DonoBoleta = 7 Then ' 35 pesos
                Pendiente.Amount = 7
            Call MeterItemEnInventario(UserIndex, Pendiente)
            UserList(UserIndex).Stats.PuntosDeTorneo = UserList(UserIndex).Stats.PuntosDeTorneo + 35
            UserList(UserIndex).Stats.Honor = UserList(UserIndex).Stats.Honor + 150
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has recibido 150 puntos de Honor." & FONTTYPE_GRISN)
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has recibido 35 puntos de Torneo." & FONTTYPE_GRISN)
        ElseIf UserList(UserIndex).Stats.DonoBoleta = 8 Then ' 40 pesos
            Pendiente.Amount = 8
            Call MeterItemEnInventario(UserIndex, Pendiente)
            UserList(UserIndex).Stats.PuntosDeTorneo = UserList(UserIndex).Stats.PuntosDeTorneo + 40
            UserList(UserIndex).Stats.Honor = UserList(UserIndex).Stats.Honor + 200
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has recibido 200 puntos de Honor." & FONTTYPE_GRISN)
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has recibido 40 puntos de Torneo." & FONTTYPE_GRISN)
        ElseIf UserList(UserIndex).Stats.DonoBoleta = 9 Then ' 45 pesos
            Pendiente.Amount = 9
            Call MeterItemEnInventario(UserIndex, Pendiente)
            UserList(UserIndex).Stats.PuntosDeTorneo = UserList(UserIndex).Stats.PuntosDeTorneo + 45
            UserList(UserIndex).Stats.Honor = UserList(UserIndex).Stats.Honor + 250
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has recibido 250 puntos de Honor." & FONTTYPE_GRISN)
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has recibido 45 puntos de Torneo." & FONTTYPE_GRISN)
        ElseIf UserList(UserIndex).Stats.DonoBoleta = 10 Then ' 50 pesos
            Pendiente.Amount = 10
            Call MeterItemEnInventario(UserIndex, Pendiente)
            UserList(UserIndex).Stats.PuntosDeTorneo = UserList(UserIndex).Stats.PuntosDeTorneo + 50
            UserList(UserIndex).Stats.Honor = UserList(UserIndex).Stats.Honor + 300
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has recibido 300 puntos de Honor." & FONTTYPE_GRISN)
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has recibido 50 puntos de Torneo." & FONTTYPE_GRISN)
        ElseIf UserList(UserIndex).Stats.DonoBoleta = 11 Then ' 60 pesos
            Pendiente.Amount = 12
            Call MeterItemEnInventario(UserIndex, Pendiente)
            UserList(UserIndex).Stats.PuntosDeTorneo = UserList(UserIndex).Stats.PuntosDeTorneo + 60
            UserList(UserIndex).Stats.Honor = UserList(UserIndex).Stats.Honor + 400
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has recibido 400 puntos de Honor." & FONTTYPE_GRISN)
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has recibido 60 puntos de Torneo." & FONTTYPE_GRISN)
        ElseIf UserList(UserIndex).Stats.DonoBoleta = 12 Then ' 70 pesos
            Pendiente.Amount = 14
            Call MeterItemEnInventario(UserIndex, Pendiente)
            UserList(UserIndex).Stats.PuntosDeTorneo = UserList(UserIndex).Stats.PuntosDeTorneo + 70
            UserList(UserIndex).Stats.Honor = UserList(UserIndex).Stats.Honor + 500
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has recibido 500 puntos de Honor." & FONTTYPE_GRISN)
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has recibido 70 puntos de Torneo." & FONTTYPE_GRISN)
        ElseIf UserList(UserIndex).Stats.DonoBoleta = 13 Then ' 80 pesos
            Pendiente.Amount = 16
            Call MeterItemEnInventario(UserIndex, Pendiente)
            UserList(UserIndex).Stats.PuntosDeTorneo = UserList(UserIndex).Stats.PuntosDeTorneo + 80
            UserList(UserIndex).Stats.Honor = UserList(UserIndex).Stats.Honor + 600
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has recibido 600 puntos de Honor." & FONTTYPE_GRISN)
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has recibido 80 puntos de Torneo." & FONTTYPE_GRISN)
        ElseIf UserList(UserIndex).Stats.DonoBoleta = 14 Then ' 90 pesos
            Pendiente.Amount = 18
            Call MeterItemEnInventario(UserIndex, Pendiente)
            UserList(UserIndex).Stats.PuntosDeTorneo = UserList(UserIndex).Stats.PuntosDeTorneo + 90
            UserList(UserIndex).Stats.Honor = UserList(UserIndex).Stats.Honor + 700
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has recibido 700 puntos de Honor." & FONTTYPE_GRISN)
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has recibido 90 puntos de Torneo." & FONTTYPE_GRISN)
        ElseIf UserList(UserIndex).Stats.DonoBoleta = 15 Then ' 100 pesos
            Pendiente.Amount = 20
            Call MeterItemEnInventario(UserIndex, Pendiente)
            UserList(UserIndex).Stats.PuntosDeTorneo = UserList(UserIndex).Stats.PuntosDeTorneo + 100
            UserList(UserIndex).Stats.Honor = UserList(UserIndex).Stats.Honor + 800
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has recibido 800 puntos de Honor." & FONTTYPE_GRISN)
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has recibido 100 puntos de Torneo." & FONTTYPE_GRISN)
        ElseIf UserList(UserIndex).Stats.DonoBoleta = 16 Then ' 110 pesos
            Pendiente.Amount = 21
            Call MeterItemEnInventario(UserIndex, Pendiente)
            UserList(UserIndex).Stats.PuntosDeTorneo = UserList(UserIndex).Stats.PuntosDeTorneo + 110
            UserList(UserIndex).Stats.Honor = UserList(UserIndex).Stats.Honor + 900
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has recibido 900 puntos de Honor." & FONTTYPE_GRISN)
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has recibido 110 puntos de Torneo." & FONTTYPE_GRISN)
        ElseIf UserList(UserIndex).Stats.DonoBoleta = 17 Then ' 120 pesos
            Pendiente.Amount = 22
            Call MeterItemEnInventario(UserIndex, Pendiente)
            UserList(UserIndex).Stats.PuntosDeTorneo = UserList(UserIndex).Stats.PuntosDeTorneo + 120
            UserList(UserIndex).Stats.Honor = UserList(UserIndex).Stats.Honor + 1000
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has recibido 100 puntos de Honor." & FONTTYPE_GRISN)
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has recibido 120 puntos de Torneo." & FONTTYPE_GRISN)
        ElseIf UserList(UserIndex).Stats.DonoBoleta = 18 Then ' 130 pesos
            Pendiente.Amount = 23
            Call MeterItemEnInventario(UserIndex, Pendiente)
            UserList(UserIndex).Stats.PuntosDeTorneo = UserList(UserIndex).Stats.PuntosDeTorneo + 130
            UserList(UserIndex).Stats.Honor = UserList(UserIndex).Stats.Honor + 1100
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has recibido 1100 puntos de Honor." & FONTTYPE_GRISN)
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has recibido 130 puntos de Torneo." & FONTTYPE_GRISN)
        ElseIf UserList(UserIndex).Stats.DonoBoleta = 19 Then ' 140 pesos
            Pendiente.Amount = 24
            Call MeterItemEnInventario(UserIndex, Pendiente)
            UserList(UserIndex).Stats.PuntosDeTorneo = UserList(UserIndex).Stats.PuntosDeTorneo + 140
            UserList(UserIndex).Stats.Honor = UserList(UserIndex).Stats.Honor + 1200
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has recibido 1200 puntos de Honor." & FONTTYPE_GRISN)
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has recibido 140 puntos de Torneo." & FONTTYPE_GRISN)
        ElseIf UserList(UserIndex).Stats.DonoBoleta = 20 Then ' 150 pesos
            Pendiente.Amount = 25
            Call MeterItemEnInventario(UserIndex, Pendiente)
            UserList(UserIndex).Stats.PuntosDeTorneo = UserList(UserIndex).Stats.PuntosDeTorneo + 150
            UserList(UserIndex).Stats.Honor = UserList(UserIndex).Stats.Honor + 1300
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has recibido 1300 puntos de Honor." & FONTTYPE_GRISN)
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has recibido 150 puntos de Torneo." & FONTTYPE_GRISN)
End If
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has recibido " & Pendiente.Amount & " Pendiente/s del sacrificio." & FONTTYPE_GRISN)
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Premio de donacion recibido correctamente, gracias por colaborar con Lhirius AO. ¡Que disfrutes!" & FONTTYPE_DAREXP)
            UserList(UserIndex).Stats.DonoBoleta = 0
            SendUserStatsBox (UserIndex)
    Exit Sub
End If
           
           If UCase$(Left$(rData, 7)) = "/DUELO " Then
 
    dMap = 5 'Mapa de duelos, cambienlo
    rData = Right$(rData, Len(rData) - 7)
    dUser = ReadField(1, rData, Asc("@"))
   
    If NameIndex(dUser) = 0 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Usuario offline." & FONTTYPE_INFO)
        Exit Sub
    Else
        dIndex = NameIndex(dUser)
    End If
   
    dMoney = ReadField(2, rData, Asc("@"))
    If dIndex = UserIndex Then
       Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No podes dueliar contra vos mismo." & FONTTYPE_INFO)
        Exit Sub
    End If
    If UserList(UserIndex).Stats.GLD < val(dMoney) Then
       Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No tenes esa cantidad de oro." & FONTTYPE_INFO)
        Exit Sub
    End If
   
    If UserList(dIndex).Stats.GLD < val(dMoney) Then
       Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El usuario no tiene esa cantidad de oro." & FONTTYPE_INFO)
        Exit Sub
    End If
     If UserList(dIndex).Pos.Map = 78 Then
       Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El usuario está encarcelado." & FONTTYPE_INFO)
        Exit Sub
    End If
    If UserList(UserIndex).flags.Muerto Then
       Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas muerto!!." & FONTTYPE_INFO)
        Exit Sub
    End If
       If UserList(UserIndex).Pos.Map = 78 Then
       Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas en la carcel." & FONTTYPE_INFO)
        Exit Sub
    End If
    If UserList(dIndex).flags.Muerto Then
       Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El usuario està muerto." & FONTTYPE_INFO)
        Exit Sub
    End If
   
    If val(dMoney) < 100000 Then
      Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El minimo de oro para duelear es de 100.000 monedas de oro." & FONTTYPE_INFO)
       Exit Sub
    End If
   
    If MapInfo(dMap).NumUsers = 2 Then
       Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Ya hay un duelo en curso, esperà que termine." & FONTTYPE_INFO)
        Exit Sub
    End If
   
    UserList(dIndex).flags.LeMandaronDuelo = True
    UserList(dIndex).flags.UltimoEnMandarDuelo = UserList(UserIndex).name
   Call SendData(SendTarget.ToIndex, (dIndex), 0, "||" & UserList(UserIndex).name & " (" & UserList(UserIndex).Clase & " - " & UserList(UserIndex).Stats.ELV & ") - te está desafiando en un duelo por " & PonerPuntos(val(dMoney)) & " monedas de oro, para aceptar escribi /SIDUELO." & "~124~124~124~1~0")
   
   Exit Sub
End If

     If UCase$(Left$(rData, 7)) = "/SIRETO" Then
   
       
        If UserList(UserIndex).flags.LeMandanReto = False Then
           Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Nadie te ha retado." & FONTTYPE_INFO)
            Exit Sub
        Else
       
        If UserList(UserIndex).flags.Muerto Then
           Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estás muerto!!." & FONTTYPE_INFO)
            Exit Sub
        End If
        If UserList(UserIndex).Pos.Map = 66 Then
           Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estás en la carcel!!." & FONTTYPE_INFO)
            Exit Sub
        End If
        
        If MapInfo(44).NumUsers = 2 Then
           Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Ya hay un reto en curso, esperá que termine." & FONTTYPE_INFO)
            Exit Sub
        End If
       
        If UserList(NameIndex(UserList(UserIndex).flags.UltimoenMandarReto)).flags.Muerto Then
           Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El usuario está muerto." & FONTTYPE_INFO)
            Exit Sub
        End If
        
        If UserList(NameIndex(UserList(UserIndex).flags.UltimoenMandarReto)).Pos.Map = 66 Then
           Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El usuario está encarcelado." & FONTTYPE_INFO)
            Exit Sub
        End If
        Dim el As Integer
        el = NameIndex(UserList(UserIndex).flags.UltimoenMandarReto)
        If Not TieneObjetos(QueOBJSeReta.ObjIndex, 1, el) Then
           Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El usuario no tiene el objeto necesario para el reto." & FONTTYPE_INFO)
            Exit Sub
        End If 'con esto evitamos que el usuario se saque el objeto del inventario antes de aceptar el reto
        
        If Not TieneObjetos(QueOBJSeReta.ObjIndex, 1, UserIndex) Then
           Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No tienes el objeto necesario para el reto." & FONTTYPE_INFO)
            Exit Sub
        End If 'con esto evitamos que el usuario se saque el objeto del inventario antes de aceptar el reto
        
        If NameIndex(UserList(UserIndex).flags.UltimoenMandarReto) = 0 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El usuario que te retó está offline." & FONTTYPE_INFO)
            Exit Sub
        End If
       
    End If
   
    
   
    UserList(el).flags.LeMandanReto = False
    UserList(el).flags.enReto = True
    UserList(UserIndex).flags.LeMandaronDuelo = False
    UserList(UserIndex).flags.enReto = True
    UserList(el).flags.RetandoContra = UserList(UserIndex).name
    UserList(UserIndex).flags.RetandoContra = UserList(el).name
    
        If ItemReto <> "" Then 'si se apuesta un item mandamos una solicitud
            SendData SendTarget.ToAll, UserIndex, 0, "||" & UserList(UserIndex).name & " y " & UserList(el).name & " van a combatir en un reto por " & FragsReto & " frags - " & PuntosReto & " puntos de torneo y 1 " & ItemReto & "." & FONTTYPE_BLANCOCN
        Call QuitarObjetos(QueOBJSeReta.ObjIndex, 1, UserIndex)
        Call QuitarObjetos(QueOBJSeReta.ObjIndex, 1, el)
        
        'mandamos la solicitud
        Else 'si no, mandamos otra solicitud
            SendData SendTarget.ToAll, UserIndex, 0, "||" & UserList(UserIndex).name & " y " & UserList(el).name & " van a combatir en un reto por " & FragsReto & " frags - " & PuntosReto & " puntos de torneo." & FONTTYPE_BLANCOCN
        'mandamos la solicitud
        End If
    UserList(UserIndex).Stats.PuntosDeTorneo = UserList(UserIndex).Stats.PuntosDeTorneo - val(PuntosReto)
    UserList(el).Stats.PuntosDeTorneo = UserList(el).Stats.PuntosDeTorneo - val(PuntosReto)
    UserList(UserIndex).Stats.UsuariosMatados = UserList(UserIndex).Stats.UsuariosMatados - val(FragsReto)
    UserList(el).Stats.UsuariosMatados = UserList(el).Stats.UsuariosMatados - val(FragsReto)
    SendUserStatsBox (UserIndex)
    SendUserStatsBox (el)
    Call WarpUserChar(el, 44, 68, 31, True)
    Call WarpUserChar(UserIndex, 44, 31, 68, True)
    Call SendUserStatsBox(UserIndex)
    Call SendUserStatsBox(el)
    
        Exit Sub
    End If
 
    If UCase$(Left$(rData, 8)) = "/SIDUELO" Then
   
       
        If UserList(UserIndex).flags.LeMandaronDuelo = False Then
           Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Nadie te ofreciò duelo." & FONTTYPE_INFO)
            Exit Sub
        Else
       
        If UserList(UserIndex).flags.Muerto Then
           Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas muerto!!." & FONTTYPE_INFO)
            Exit Sub
        End If
        If UserList(UserIndex).Pos.Map = 66 Then
           Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas en la carcel!!." & FONTTYPE_INFO)
            Exit Sub
        End If
        If UserList(UserIndex).Stats.GLD < val(dMoney) Then
           Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No tenes " & PonerPuntos(val(dMoney)) & " monedas de oro para aceptar el duelo." & FONTTYPE_INFO)
            Exit Sub
        End If
     
        If MapInfo(43).NumUsers = 2 Then
           Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Ya hay un duelo en curso, esperà que termine." & FONTTYPE_INFO)
            Exit Sub
        End If
       
        If UserList(NameIndex(UserList(UserIndex).flags.UltimoEnMandarDuelo)).flags.Muerto Then
           Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El usuario està muerto." & FONTTYPE_INFO)
            Exit Sub
        End If
        If UserList(NameIndex(UserList(UserIndex).flags.UltimoEnMandarDuelo)).Pos.Map = 66 Then
           Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El usuario está encarcelado." & FONTTYPE_INFO)
            Exit Sub
        End If
        If UserList(NameIndex(UserList(UserIndex).flags.UltimoEnMandarDuelo)).Stats.GLD < val(dMoney) Then
           Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El usuario no tiene el oro suficiente para hacer el duelo." & FONTTYPE_INFO)
            Exit Sub
        End If
       
        If NameIndex(UserList(UserIndex).flags.UltimoEnMandarDuelo) = 0 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El usuario que te mandò duelo, està offline." & FONTTYPE_INFO)
            Exit Sub
        End If
       
    End If
   
    
    el = NameIndex(UserList(UserIndex).flags.UltimoEnMandarDuelo)
   
    UserList(el).flags.LeMandaronDuelo = False
    UserList(el).flags.Endueloo = True
    UserList(UserIndex).flags.LeMandaronDuelo = False
    UserList(UserIndex).flags.Endueloo = True
    UserList(el).flags.DueliandoContra = UserList(UserIndex).name
    UserList(UserIndex).flags.DueliandoContra = UserList(el).name
    SendData SendTarget.ToAll, UserIndex, 0, "||" & UserList(UserIndex).name & " y " & UserList(el).name & " van a combatir en un duelo por " & PonerPuntos(val(dMoney)) & " monedas de oro." & FONTTYPE_TALK
 
    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - val(dMoney)
    UserList(el).Stats.GLD = UserList(el).Stats.GLD - val(dMoney)
    SendUserStatsBox (UserIndex)
    SendUserStatsBox (el)
    Call WarpUserChar(el, 43, 36, 59, True)
    Call WarpUserChar(UserIndex, 43, 64, 41, True)
    Call SendUserStatsBox(UserIndex)
    Call SendUserStatsBox(el)
        Exit Sub
    End If
    
    If UCase$(Left$(rData, 6)) = "/PMSG " Then
        If Len(rData) > 6 Then
            Call mdParty.BroadCastParty(UserIndex, mid$(rData, 7))
            Call SendData(SendTarget.ToPartyArea, UserIndex, UserList(UserIndex).Pos.Map, "||" & vbYellow & "°< " & mid$(rData, 7) & " >°" & CStr(UserList(UserIndex).Char.CharIndex))
        End If
        Exit Sub
    End If
   'Pos clan
    If UCase$(rData) = "/CLAN" Then
        tStr = modGuilds.m_ListaDeMiembrosOnlinePOS(UserIndex, UserList(UserIndex).GuildIndex)
        If UserList(UserIndex).GuildIndex <> 0 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Miembros: " & tStr & FONTTYPE_GUILDMSG)
        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No pertences a ningún clan." & FONTTYPE_GUILDMSG)
        End If
        Exit Sub
    End If
    If UCase$(Left$(rData, 11)) = "/CENTINELA " Then
        'Evitamos overflow y underflow
        If val(Right$(rData, Len(rData) - 11)) > &H7FFF Or val(Right$(rData, Len(rData) - 11)) < &H8000 Then Exit Sub
        
        tInt = val(Right$(rData, Len(rData) - 11))
        Call CentinelaCheckClave(UserIndex, tInt)
        Exit Sub
    End If
    
    If UCase$(rData) = "/ONLINECLAN" Then
        tStr = modGuilds.m_ListaDeMiembrosOnline(UserIndex, UserList(UserIndex).GuildIndex)
        If UserList(UserIndex).GuildIndex <> 0 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Compañeros de tu clan conectados: " & tStr & FONTTYPE_GUILDMSG)
        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No pertences a ningún clan." & FONTTYPE_GUILDMSG)
        End If
        Exit Sub
    End If
    If UCase$(Left$(rData, 6)) = "/NICK " Then
rData = Right$(rData, Len(rData) - 6)
 
name = rData
tIndex = NameIndex(name)
If tIndex <= 0 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El usuario NO esta online." & FONTTYPE_FIGHT)
        Exit Sub
    End If
 
If tIndex <= 1 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El usuario " & UserList(UserIndex).name & " SI esta online." & FONTTYPE_SERVER)
        Exit Sub
    End If
End If
    If UCase$(rData) = "/ONLINEPARTY" Then
        Call mdParty.OnlineParty(UserIndex)
        Exit Sub
    End If
   If UCase$(Left$(rData, 8)) = "/DARORO " Then
        Dim Cantidad As Long
        Cantidad = UserList(UserIndex).Stats.GLD
        Call LogGM(UserList(UserIndex).name, rData, False)
        rData = Right$(rData, Len(rData) - 8)
        tIndex = NameIndex(ReadField(1, rData, 32))
        Arg1 = ReadField(2, rData, 32)
        If tIndex <= 0 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Usuario offline." & FONTTYPE_INFO)
            Exit Sub
        End If
        If val(Arg1) > Cantidad Then
            Call SendUserStatsBox(tIndex)
            Call SendUserStatsBox(UserIndex)
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No tenes esa cantidad de oro" & FONTTYPE_WARNING)
        ElseIf val(Arg1) < 0 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No podes transferir cantidades negativas" & FONTTYPE_WARNING)
            Call SendUserStatsBox(tIndex)
            Call SendUserStatsBox(UserIndex)
        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡Le regalaste " & val(Arg1) & " monedas de oro a " & UserList(tIndex).name & "!" & FONTTYPE_WARNING)
            Call SendData(SendTarget.ToIndex, tIndex, 0, "||¡" & UserList(UserIndex).name & " te regalo " & val(Arg1) & " monedas de oro!" & FONTTYPE_WARNING)
            UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - val(Arg1)
            UserList(tIndex).Stats.GLD = UserList(tIndex).Stats.GLD + val(Arg1)
            Call SendUserStatsBox(tIndex)
            Call SendUserStatsBox(UserIndex)
            Exit Sub
        End If
        Exit Sub
    End If
    '[yb]
    If UCase$(Left$(rData, 6)) = "/BMSG " Then
        rData = Right$(rData, Len(rData) - 6)
        If UserList(UserIndex).flags.PertAlCons = 1 Then
            Call SendData(SendTarget.ToConsejo, UserIndex, 0, "|| (Consejero) " & UserList(UserIndex).name & "> " & rData & FONTTYPE_CONSEJO)
        End If
        If UserList(UserIndex).flags.PertAlConsCaos = 1 Then
            Call SendData(SendTarget.ToConsejoCaos, UserIndex, 0, "|| (Consejero) " & UserList(UserIndex).name & "> " & rData & FONTTYPE_CONSEJOCAOS)
        End If
        Exit Sub
    End If
    '[/yb]
    
    If UCase$(Left$(rData, 5)) = "/ROL " Then
        rData = Right$(rData, Len(rData) - 5)
        Call SendData(SendTarget.ToIndex, 0, 0, "|| " & "Su solicitud ha sido enviada" & FONTTYPE_INFO)
        Call SendData(SendTarget.ToRolesMasters, 0, 0, "|| " & LCase$(UserList(UserIndex).name) & " PREGUNTA ROL: " & rData & FONTTYPE_GUILDMSG)
        Exit Sub
    End If
    
    
    'Mensaje del servidor a GMs - Lo ubico aqui para que no se confunda con /GM [Gonzalo]
    If UCase$(Left$(rData, 6)) = "/GMSG " And UserList(UserIndex).flags.Privilegios > PlayerType.User Then
        rData = Right$(rData, Len(rData) - 6)
        Call LogGM(UserList(UserIndex).name, "Mensaje a Gms:" & rData, (UserList(UserIndex).flags.Privilegios = PlayerType.Consejero))
        If rData <> "" Then
            Call SendData(SendTarget.ToAdmins, 0, 0, "||" & UserList(UserIndex).name & "> " & rData & "~255~255~255~0~1")
        End If
        Exit Sub
    End If
    
    Select Case UCase$(Left$(rData, 3))
        Case "/GM"
            If Not Ayuda.Existe(UserList(UserIndex).name) Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El mensaje ha sido entregado, ahora solo debes esperar que se desocupe algun GM." & FONTTYPE_INFO)
                Call Ayuda.Push(rData, UserList(UserIndex).name)
            Else
                Call Ayuda.Quitar(UserList(UserIndex).name)
                Call Ayuda.Push(rData, UserList(UserIndex).name)
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Ya habias mandado un mensaje, tu mensaje ha sido movido al final de la cola de mensajes." & FONTTYPE_INFO)
            End If
            Exit Sub
    End Select
    
    
    
    Select Case UCase(Left(rData, 5))
        Case "/_BUG "
            N = FreeFile
            Open App.Path & "\LOGS\BUGs.log" For Append Shared As N
            Print #N,
            Print #N,
            Print #N, "########################################################################"
            Print #N, "########################################################################"
            Print #N, "Usuario:" & UserList(UserIndex).name & "  Fecha:" & Date & "    Hora:" & Time
            Print #N, "########################################################################"
            Print #N, "BUG:"
            Print #N, Right$(rData, Len(rData) - 5)
            Print #N, "########################################################################"
            Print #N, "########################################################################"
            Print #N,
            Print #N,
            Close #N
            Exit Sub
    
    End Select
    
    Select Case UCase$(Left$(rData, 6))
        Case "/DESC "
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes cambiar la descripción estando muerto." & FONTTYPE_INFO)
                Exit Sub
            End If
            rData = Right$(rData, Len(rData) - 6)
            If Not AsciiValidos(rData) Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||La descripcion tiene caracteres invalidos." & FONTTYPE_INFO)
                Exit Sub
            End If
            UserList(UserIndex).Desc = Trim$(rData)
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||La descripcion a cambiado." & FONTTYPE_INFO)
            Exit Sub
        Case "/VOTO "
                rData = Right$(rData, Len(rData) - 6)
                If Not modGuilds.v_UsuarioVota(UserIndex, rData, tStr) Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Voto NO contabilizado: " & tStr & FONTTYPE_GUILD)
                Else
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Voto contabilizado." & FONTTYPE_GUILD)
                End If
                Exit Sub
    End Select
    
    If UCase$(Left$(rData, 7)) = "/PENAS " Then
        name = Right$(rData, Len(rData) - 7)
        If name = "" Then Exit Sub
        
        name = Replace(name, "\", "")
        name = Replace(name, "/", "")
        
        If FileExist(CharPath & name & ".chr", vbNormal) Then
            tInt = val(GetVar(CharPath & name & ".chr", "PENAS", "Cant"))
            If tInt = 0 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Sin prontuario.." & FONTTYPE_INFO)
            Else
                While tInt > 0
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & tInt & "- " & GetVar(CharPath & name & ".chr", "PENAS", "P" & tInt) & FONTTYPE_INFO)
                    tInt = tInt - 1
                Wend
            End If
        Else
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Personaje """ & name & """ inexistente." & FONTTYPE_INFO)
        End If
        Exit Sub
    End If
    
    
    
    
    
    Select Case UCase$(Left$(rData, 8))
        Case "/PASSWD "
            rData = Right$(rData, Len(rData) - 8)
            If Len(rData) < 6 Then
                 Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El password debe tener al menos 6 caracteres." & FONTTYPE_INFO)
            Else
                 Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El password ha sido cambiado." & FONTTYPE_INFO)
                 UserList(UserIndex).Password = rData
            End If
            Exit Sub
    End Select
    
    Select Case UCase$(Left$(rData, 9))
            'Comando /APOSTAR basado en la idea de DarkLight,
            'pero con distinta probabilidad de exito.
        Case "/APOSTAR "
            rData = Right(rData, Len(rData) - 9)
            tLong = CLng(val(rData))
            If tLong > 32000 Then tLong = 32000
            N = tLong
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
            ElseIf UserList(UserIndex).flags.TargetNPC = 0 Then
                'Se asegura que el target es un npc
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
            ElseIf Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 10 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
            ElseIf Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> eNPCType.Timbero Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "No tengo ningun interes en apostar." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
            ElseIf N < 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "El minimo de apuesta es 1 moneda." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
            ElseIf N > 5000 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "El maximo de apuesta es 5000 monedas." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
            ElseIf UserList(UserIndex).Stats.GLD < N Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "No tienes esa cantidad." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
            Else
                If RandomNumber(1, 100) <= 47 Then
                    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + N
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Felicidades! Has ganado " & CStr(N) & " monedas de oro!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
                    
                    Apuestas.Perdidas = Apuestas.Perdidas + N
                    Call WriteVar(DatPath & "apuestas.dat", "Main", "Perdidas", CStr(Apuestas.Perdidas))
                Else
                    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - N
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Lo siento, has perdido " & CStr(N) & " monedas de oro." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
                
                    Apuestas.Ganancias = Apuestas.Ganancias + N
                    Call WriteVar(DatPath & "apuestas.dat", "Main", "Ganancias", CStr(Apuestas.Ganancias))
                End If
                Apuestas.Jugadas = Apuestas.Jugadas + 1
                Call WriteVar(DatPath & "apuestas.dat", "Main", "Jugadas", CStr(Apuestas.Jugadas))
                
                Call SendUserStatsBox(UserIndex)
            End If
            Exit Sub
    End Select
    
    Select Case UCase$(Left$(rData, 10))
            'consultas populares muchacho'
        Case "/ENCUESTA "
            rData = Right(rData, Len(rData) - 10)
            If Len(rData) = 0 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "|| Aca va la info de la encuesta" & FONTTYPE_GUILD)
                Exit Sub
            End If
            DummyInt = CLng(val(rData))
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "|| " & ConsultaPopular.doVotar(UserIndex, DummyInt) & FONTTYPE_GUILD)
            Exit Sub
    End Select
 
    Select Case UCase$(Left$(rData, 8))
        Case "/RETIRAR" 'RETIRA ORO EN EL BANCO o te saca de la armada
             '¿Esta el user muerto? Si es asi no puede comerciar
             If UserList(UserIndex).flags.Muerto = 1 Then
                      Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                      Exit Sub
             End If
             'Se asegura que el target es un npc
             If UserList(UserIndex).flags.TargetNPC = 0 Then
                  Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
                  Exit Sub
             End If
             
             If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype = 5 Then
                
                'Se quiere retirar de la armada
                If UserList(UserIndex).Faccion.ArmadaReal = 1 Then
                    If Npclist(UserList(UserIndex).flags.TargetNPC).flags.Faccion = 0 Then
                        Call ExpulsarFaccionReal(UserIndex)
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "º" & "Serás bienvenido a las fuerzas imperiales si deseas regresar." & "º" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
                        Debug.Print "||" & vbWhite & "º" & "Serás bienvenido a las fuerzas imperiales si deseas regresar." & "º" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex)
                    Else
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "º" & "¡¡¡Sal de aquí bufón!!!" & "º" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
                    End If
                ElseIf UserList(UserIndex).Faccion.FuerzasCaos = 1 Then
                    If Npclist(UserList(UserIndex).flags.TargetNPC).flags.Faccion = 1 Then
                        Call ExpulsarFaccionCaos(UserIndex)
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "º" & "Ya volverás arrastrandote." & "º" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
                    Else
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "º" & "Sal de aquí maldito criminal" & "º" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
                    End If
                Else
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "º" & "¡No perteneces a ninguna fuerza!" & "º" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
                End If
                Exit Sub
             
             End If
             
             If Len(rData) = 8 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Debes indicar el monto de cuanto quieres retirar" & FONTTYPE_INFO)
                Exit Sub
             End If
             
             rData = Right$(rData, Len(rData) - 9)
             If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> eNPCType.Banquero _
             Or UserList(UserIndex).flags.Muerto = 1 Then Exit Sub
             If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNPC).Pos) > 10 Then
                  Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                  Exit Sub
             End If
             If FileExist(CharPath & UCase$(UserList(UserIndex).name) & ".chr", vbNormal) = False Then
                  Call SendData(SendTarget.ToIndex, UserIndex, 0, "!!El personaje no existe, cree uno nuevo.")
                  CloseSocket (UserIndex)
                  Exit Sub
             End If
             If val(rData) > 0 And val(rData) <= UserList(UserIndex).Stats.Banco Then
                  UserList(UserIndex).Stats.Banco = UserList(UserIndex).Stats.Banco - val(rData)
                  UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD + val(rData)
                  Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Tenes " & UserList(UserIndex).Stats.Banco & " monedas de oro en tu cuenta." & "°" & Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex & FONTTYPE_INFO)
             Else
                  Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & " No tenes esa cantidad." & "°" & Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex & FONTTYPE_INFO)
             End If
             Call SendUserStatsBox(val(UserIndex))
             Exit Sub
    End Select
    
    Select Case UCase$(Left$(rData, 11))
        Case "/DEPOSITAR " 'DEPOSITAR ORO EN EL BANCO
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(UserIndex).flags.Muerto = 1 Then
                      Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                      Exit Sub
            End If
            'Se asegura que el target es un npc
            If UserList(UserIndex).flags.TargetNPC = 0 Then
                  Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Primero tenes que seleccionar un personaje, hace click izquierdo sobre el." & FONTTYPE_INFO)
                  Exit Sub
            End If
            If Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 10 Then
                      Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                      Exit Sub
            End If
            rData = Right$(rData, Len(rData) - 11)
            If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> eNPCType.Banquero _
            Or UserList(UserIndex).flags.Muerto = 1 Then Exit Sub
            If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNPC).Pos) > 10 Then
                  Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                  Exit Sub
            End If
            If CLng(val(rData)) > 0 And CLng(val(rData)) <= UserList(UserIndex).Stats.GLD Then
                  UserList(UserIndex).Stats.Banco = UserList(UserIndex).Stats.Banco + val(rData)
                  UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - val(rData)
                  Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & "Tenes " & UserList(UserIndex).Stats.Banco & " monedas de oro en tu cuenta." & "°" & Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex & FONTTYPE_INFO)
            Else
                  Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & vbWhite & "°" & " No tenes esa cantidad." & "°" & Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex & FONTTYPE_INFO)
            End If
            Call SendUserStatsBox(val(UserIndex))
            Exit Sub
        Case "/DENUNCIAR "
            If UserList(UserIndex).flags.Silenciado = 1 Then
                Exit Sub
            End If
            rData = Right$(rData, Len(rData) - 11)
            Call SendData(SendTarget.ToAdmins, 0, 0, "|| " & LCase$(UserList(UserIndex).name) & " DENUNCIA: " & rData & FONTTYPE_GUILDMSG)
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "|| Denuncia enviada, espere.." & FONTTYPE_INFO)
            Exit Sub
        Case "/FUNDARCLAN"
        
            rData = Right$(rData, Len(rData) - 11)
            If Trim$(rData) = vbNullString Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "|| Para fundar un clan debes especificar la alineación del mismo." & FONTTYPE_GUILD)
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "|| Atención, que la misma no podrá cambiar luego, te aconsejamos leer las reglas sobre clanes antes de fundar." & FONTTYPE_GUILD)
                Exit Sub
            Else
                Select Case UCase$(Trim(rData))
                    Case "0"
                        UserList(UserIndex).FundandoGuildAlineacion = ALINEACION_NEUTRO
                    Case "GM"
                        UserList(UserIndex).FundandoGuildAlineacion = ALINEACION_MASTER
                    Case "2"
                        UserList(UserIndex).FundandoGuildAlineacion = ALINEACION_CIUDA
                    Case "1"
                        UserList(UserIndex).FundandoGuildAlineacion = ALINEACION_CRIMINAL
                    Case Else
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "|| Alineación inválida." & FONTTYPE_GUILD)
                        Exit Sub
                End Select
            End If

            If modGuilds.PuedeFundarUnClan(UserIndex, UserList(UserIndex).FundandoGuildAlineacion, tStr) Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "SHOWFUN")
            Else
                UserList(UserIndex).FundandoGuildAlineacion = 0
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & tStr & FONTTYPE_GUILD)
            End If
            
            Exit Sub
    
    End Select

    Select Case UCase$(Left$(rData, 12))
        Case "/ECHARPARTY "
            rData = Right$(rData, Len(rData) - 12)
            tInt = NameIndex(rData)
            If tInt > 0 Then
                Call mdParty.ExpulsarDeParty(UserIndex, tInt)
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "|| El personaje no está online." & FONTTYPE_INFO)
            End If
            Exit Sub
        Case "/PARTYLIDER "
            rData = Right$(rData, Len(rData) - 12)
            tInt = NameIndex(rData)
            If tInt > 0 Then
                Call mdParty.TransformarEnLider(UserIndex, tInt)
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "|| El personaje no está online." & FONTTYPE_INFO)
            End If
            Exit Sub
    
    End Select

    Select Case UCase$(Left$(rData, 13))
        Case "/ACCEPTPARTY "
            rData = Right$(rData, Len(rData) - 13)
            tInt = NameIndex(rData)
            If tInt > 0 Then
                Call mdParty.AprobarIngresoAParty(UserIndex, tInt)
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "|| El personaje no está online." & FONTTYPE_INFO)
            End If
            Exit Sub
    
    End Select
    

    Select Case UCase$(Left$(rData, 14))
        Case "/MIEMBROSCLAN "
            rData = Trim(Right(rData, Len(rData) - 14))
            name = Replace(rData, "\", "")
            name = Replace(rData, "/", "")
    
            If Not FileExist(App.Path & "\guilds\" & rData & "-members.mem") Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "|| No existe el clan: " & rData & FONTTYPE_INFO)
                Exit Sub
            End If
            
            tInt = val(GetVar(App.Path & "\Guilds\" & rData & "-Members" & ".mem", "INIT", "NroMembers"))
            
            For i = 1 To tInt
                tStr = GetVar(App.Path & "\Guilds\" & rData & "-Members" & ".mem", "Members", "Member" & i)
                'tstr es la victima
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & tStr & "<" & rData & ">." & FONTTYPE_INFO)
            Next i
        
            Exit Sub
    End Select
    
    Procesado = False
End Sub
