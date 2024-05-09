Attribute VB_Name = "TCP_HandleData1"
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

Public Sub HandleData_1(ByVal UserIndex As Integer, rData As String, ByRef Procesado As Boolean)


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

    Select Case UCase$(Left$(rData, 11))
        'Retos by Dylan.-
    Case "PREPARORETO"
    rData = Right$(rData, Len(rData) - 11)

        'leemos lo que mando el usuario
    nickReto = ReadField(2, rData, 44)
    PuntosReto = ReadField(3, rData, 44)
    FragsReto = ReadField(4, rData, 44)
    ItemReto = ReadField(5, rData, 44)
    
    'empezamos las condiciones para retar....
    If NameIndex(nickReto) = 0 Then 'si esta online
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Usuario offline." & FONTTYPE_INFO)
        Exit Sub
    Else
        nickReto = NameIndex(nickReto)
    End If
    
    If nickReto = UserIndex Then 'si es el mismo usuario
       Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No podes retar contra vos mismo." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    If UserList(UserIndex).flags.Muerto Then 'si ta muerto
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estás muerto." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    If UserList(nickReto).flags.Muerto Then 'si el oponente ta muerto
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Tu oponente está muerto." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    If UserList(UserIndex).Pos.Map = 66 Or UserList(nickReto).Pos.Map = 66 Then 'si uno de los 2 users esta en carcel
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Uno de ustedes dos está en la carcel." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    If UserList(UserIndex).flags.EnCvc = True Or UserList(UserIndex).flags.Endueloo Or UserList(UserIndex).flags.EstoyEnTorneo = True Then 'si esta en cvc o duelo or torneo
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes retar estando acá." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    If UserList(nickReto).flags.EnCvc = True Or UserList(nickReto).flags.Endueloo Or UserList(nickReto).flags.EstoyEnTorneo = True Then 'si esta en cvc o duelo o torneo
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El otro usuario está en duelos." & FONTTYPE_INFO)
        Exit Sub
    End If
    If UCase$(PuntosReto) <> "PUNTOS DE TORNEO" Then 'si apuesta por puntos
        If PuntosReto < 30 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Tienes que retar por mas de 30 puntos de torneo." & FONTTYPE_INFO)
            Exit Sub
        End If
    End If
    If UCase$(PuntosReto) <> "PUNTOS DE TORNEO" Then 'si apuesta por puntos
    If UserList(UserIndex).Stats.PuntosDeTorneo < PuntosReto Then 'si no tiene suficientes puntos
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No tienes " & PuntosReto & " puntos de torneo para retar." & FONTTYPE_INFO)
            Exit Sub
    End If
    End If
    
    If UCase$(PuntosReto) <> "PUNTOS DE TORNEO" Then 'si apuesta por puntos
    If UserList(nickReto).Stats.PuntosDeTorneo < PuntosReto Then 'si el oponente no tiene suficientes puntos
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Tu oponente no tiene " & PuntosReto & " puntos de torneo para aceptar tu reto." & FONTTYPE_INFO)
            Exit Sub
    End If
    End If
    
    If UCase$(FragsReto) <> "FRAGS" Then 'si apuesta por frags
        If FragsReto < 4 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Tienes que retar por mas de 4 Frags." & FONTTYPE_INFO)
            Exit Sub
        End If
    End If
    
    If UCase$(FragsReto) <> "FRAGS" Then 'si apuesta por frags
    If UserList(UserIndex).Stats.UsuariosMatados < FragsReto Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No tienes " & FragsReto & " Frags para retar." & FONTTYPE_INFO)
            Exit Sub
    End If
    End If
    
    If UCase$(FragsReto) <> "FRAGS" Then 'si apuesta por frags
    If UserList(nickReto).Stats.UsuariosMatados < FragsReto Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Tu oponente no tiene " & FragsReto & " Frags para aceptar tu reto." & FONTTYPE_INFO)
            Exit Sub
    End If
    End If
    
    If UCase$(PuntosReto) = "PUNTOS DE TORNEO" And UCase$(FragsReto) = "FRAGS" And ItemReto = "" Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Debes retar al menos por una opción." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    
    
    QueOBJSeReta.Amount = 1
    
    If UCase$(ItemReto) = "DAGA DE HIELO" Then
        QueOBJSeReta.ObjIndex = 854
    ElseIf UCase$(ItemReto) = "ESPADA DE LAS ALMAS" Then
        QueOBJSeReta.ObjIndex = 856
    ElseIf UCase$(ItemReto) = "ARMADURA HEROICA" Then
        QueOBJSeReta.ObjIndex = 857
    ElseIf UCase$(ItemReto) = "ARMADURA DE CAMPEÓN" Then
        QueOBJSeReta.ObjIndex = 858
    ElseIf UCase$(ItemReto) = "ESCUDO DE DRAGÓN" Then
        QueOBJSeReta.ObjIndex = 859
    ElseIf UCase$(ItemReto) = "ESPADA ARGENTUM" Then
        QueOBJSeReta.ObjIndex = 860
    ElseIf UCase$(ItemReto) = "CASCO OSCURO" Then
        QueOBJSeReta.ObjIndex = 861
    ElseIf UCase$(ItemReto) = "CORONA" Then
        QueOBJSeReta.ObjIndex = 862
    ElseIf UCase$(ItemReto) = "ESPADA DEL BARLOG" Then
        QueOBJSeReta.ObjIndex = 863
    ElseIf UCase$(ItemReto) = "DAGA INFERNAL" Then
        QueOBJSeReta.ObjIndex = 864
    ElseIf UCase$(ItemReto) = "CETRO DE ARCHIMAGO" Then
        QueOBJSeReta.ObjIndex = 865
    ElseIf UCase$(ItemReto) = "ANILLO DE LOS DIOSES" Then
        QueOBJSeReta.ObjIndex = 866
    ElseIf UCase$(ItemReto) = "ARMADURA DE DRAGÓN OSCURO (ALTOS)" Then
        QueOBJSeReta.ObjIndex = 867
    ElseIf UCase$(ItemReto) = "ARMADURA DE DRAGÓN OSCURO (BAJOS)" Then
        QueOBJSeReta.ObjIndex = 868
    ElseIf UCase$(ItemReto) = "ARCO ARGENTUM" Then
        QueOBJSeReta.ObjIndex = 869
    ElseIf UCase$(ItemReto) = "TÚNICA DE APOCALIPSIS (ALTOS)" Then
        QueOBJSeReta.ObjIndex = 870
    ElseIf UCase$(ItemReto) = "TÚNICA DE APOCALIPSIS (ALTOS)" Then
        QueOBJSeReta.ObjIndex = 871
    ElseIf UCase$(ItemReto) = "MANTO ALADO (ALTOS)" Then
        QueOBJSeReta.ObjIndex = 872
    ElseIf UCase$(ItemReto) = "MANTO ALADO (BAJOS)" Then
        QueOBJSeReta.ObjIndex = 873
    ElseIf UCase$(ItemReto) = "AMULETO DEL LIDER" Then
        QueOBJSeReta.ObjIndex = 874
    ElseIf UCase$(ItemReto) = "PENDIENTE DEL SACRIFICIO" Then
        QueOBJSeReta.ObjIndex = 875
    ElseIf UCase$(ItemReto) = "INSURRECCIÓN SOMBRIA" Then
        QueOBJSeReta.ObjIndex = 855
    ElseIf UCase$(ItemReto) = "PRISIÓN GELIDA" Then
        QueOBJSeReta.ObjIndex = 855
    ElseIf UCase$(ItemReto) = "FUEGO DIVINO" Then
        QueOBJSeReta.ObjIndex = 855
    ElseIf UCase$(ItemReto) = "INFIERNO" Then
        QueOBJSeReta.ObjIndex = 855
    ElseIf UCase$(ItemReto) = "VIENTO HURACANADO" Then
        QueOBJSeReta.ObjIndex = 855
        End If
        
    'si el que es retado tiene el objeto que se reta
    If Not TieneObjetos(QueOBJSeReta.ObjIndex, 1, nickReto) Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Tu oponente no tiene el objeto que ustéd desea apostar en este reto." & FONTTYPE_INFO)
        Exit Sub
    End If
    
    If MapInfo(2).NumUsers > 2 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Hay un reto en curso, intenta mas tarde." & FONTTYPE_INFO)
        Exit Sub
    End If
'---------------------------------------------------FIN    CONDICIONES     (SIST. RETOS) --------------------------------------------------------------------
        If UCase$(PuntosReto) = "PUNTOS DE TORNEO" Then
        PuntosReto = 0
        End If
        
        If UCase$(FragsReto) = "FRAGS" Then
        FragsReto = 0
        End If
        
        If ItemReto <> "" Then 'si se apuesta un item mandamos una solicitud
           Call SendData(SendTarget.ToIndex, (nickReto), 0, "||" & UserList(UserIndex).name & " (" & UserList(UserIndex).Clase & " - " & UserList(UserIndex).Stats.ELV & ") - te está desafiando en un reto por " & FragsReto & " frags - " & PuntosReto & " puntos de torneo y 1 " & ItemReto & ", para aceptar escribi /SIRETO." & FONTTYPE_BLANCOCN)
        'mandamos la solicitud
    UserList(nickReto).flags.LeMandanReto = True
    UserList(nickReto).flags.UltimoenMandarReto = UserList(UserIndex).name
    
        Else 'si no, mandamos otra solicitud
            Call SendData(SendTarget.ToIndex, (nickReto), 0, "||" & UserList(UserIndex).name & " (" & UserList(UserIndex).Clase & " - " & UserList(UserIndex).Stats.ELV & ") - te está desafiando en un reto por " & FragsReto & " frags - " & PuntosReto & " puntos de torneo, para aceptar escribi /SIRETO." & FONTTYPE_BLANCOCN)
        'mandamos la solicitud
    UserList(nickReto).flags.LeMandanReto = True
    UserList(nickReto).flags.UltimoenMandarReto = UserList(UserIndex).name
        End If
            Exit Sub
    End Select
    
    Select Case UCase$(Left$(rData, 3))
        'Información de los objetos
    Case "IPX"
        rData = Right$(rData, Len(rData) - 3)
           
            If val(rData) > 0 And val(rData) < UBound(PremiosList) + 1 Then _
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "INF" & PremiosList(val(rData)).ObjRequiere & "," & PremiosList(val(rData)).ObjMaxAt & "," & PremiosList(val(rData)).ObjMinAt & "," & PremiosList(val(rData)).ObjMaxdef & "," & PremiosList(val(rData)).ObjMindef & "," & PremiosList(val(rData)).ObjMaxAtMag & "," & PremiosList(val(rData)).ObjMinAtMag & "," & PremiosList(val(rData)).ObjMaxDefMag & "," & PremiosList(val(rData)).ObjMinDefMag & "," & PremiosList(rData).ObjDescripcion & "," & UserList(UserIndex).Stats.PuntosDeTorneo & "," & ObjData(PremiosList(rData).ObjIndexP).GrhIndex)
        Exit Sub
    
    'Requerimientos de los objetos
    Case "SPX"
        rData = Right$(rData, Len(rData) - 3)
        Dim premio As Obj
           
            If val(rData) > 0 And val(rData) < UBound(PremiosList) + 1 Then
     
            premio.Amount = 1
            premio.ObjIndex = PremiosList(val(rData)).ObjIndexP
            
            End If
           
            'Si no tiene los puntos necesarios
            If UserList(UserIndex).Stats.PuntosDeTorneo < PremiosList(val(rData)).ObjRequiere Then
                   Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No tienes suficientes puntos para este objeto." & FONTTYPE_INFO)
            Exit Sub
            End If
           
            'Si no tenemoss lugar lo tiramos al piso
            'If Not MeterItemEnInventario(UserIndex, Premio) Then
            '   Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedo cargar mas objetos." & FONTTYPE_INFO)
            'Exit Sub
            'End If
           
            'Metemos en inventario
            Call MeterItemEnInventario(UserIndex, premio)
            Call UpdateUserInv(True, UserIndex, 0)
       
            'Avisamos por consola
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has obtenido: " & ObjData(premio.ObjIndex).name & " (Cantidad: " & premio.Amount & ")" & FONTTYPE_GUILD)
           
            'Restamos & actualizams
            UserList(UserIndex).Stats.PuntosDeTorneo = UserList(UserIndex).Stats.PuntosDeTorneo - PremiosList(val(rData)).ObjRequiere
            Call SendUserStatsBox(UserIndex)
        Exit Sub
    End Select
    'Dylan - Sistema de Premios
    Select Case UCase$(Left$(rData, 1))
          Case "X"        ' >>> Sistema Consultas - Fishar.-
            rData = Right$(rData, Len(rData) - 1)
            Dim Usuario As Integer
            Dim texto As String
            Usuario = NameIndex(ReadField(1, rData, Asc("*")))
            texto = ReadField(2, rData, Asc("*"))
            Call SendData(SendTarget.ToIndex, Usuario, 0, "||Tu pregunta ah sido respondida, para verla, escribì /GM y clickea el boton 'Respuesta'." & "~255~255~255~1~0")
            Call SendData(SendTarget.ToIndex, Usuario, 0, "RESPUES" & texto)
            Exit Sub
        Case "#"       ' >>> Sistema Consultas - Fishar.-
        Debug.Print "Me llego SOS"
            rData = Right$(rData, Len(rData) - 1)
            Dim TipoConsulta As Byte
            Dim rDatax As String
            TipoConsulta = ReadField(1, rData, Asc(","))
            rDatax = ReadField(2, rData, Asc(","))
   
            If UserList(UserIndex).flags.Silenciado = True Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estás silenciado." & FONTTYPE_INFO)
                Exit Sub
            End If
       
            If TipoConsulta = 0 Then
                If Not Ayuda.Existe(UserList(UserIndex).name) Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El mensaje ha sido entregado, ahora solo debes esperar que se desocupe algun GM." & FONTTYPE_INFO)
                    Call SendData(SendTarget.ToAdmins, 0, 0, "||Nuevo mensaje S.o.S" & "~195~250~35~1~0")
                    Call SendData(SendTarget.ToAdmins, 0, 0, "NEWSOSM[Pregunta]," & UserList(UserIndex).name & "," & rDatax)
                    Call Ayuda.Push(rData, UserList(UserIndex).name)
                Else
                    Call Ayuda.Quitar(UserList(UserIndex).name)
                    Call Ayuda.Push(rData, UserList(UserIndex).name)
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Ya habias mandado un mensaje, tu mensaje ha sido movido al final de la cola de mensajes." & FONTTYPE_INFO)
                End If
            ElseIf TipoConsulta = 1 Then
                If Not Ayuda.Existe(UserList(UserIndex).name) Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El mensaje ha sido entregado, ahora solo debes esperar que se desocupe algun GM." & FONTTYPE_INFO)
                    Call SendData(SendTarget.ToAdmins, 0, 0, "||Nuevo mensaje S.o.S" & "~195~250~35~1~0")
                    Call SendData(SendTarget.ToAdmins, 0, 0, "NEWSOSM[Descargo]," & UserList(UserIndex).name & "," & rDatax)
                    Call Ayuda.Push(rData, UserList(UserIndex).name & "++")
                Else
                    Call Ayuda.Quitar(UserList(UserIndex).name)
                    Call Ayuda.Push(rData, UserList(UserIndex).name)
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Ya habias mandado un mensaje, tu mensaje ha sido movido al final de la cola de mensajes." & FONTTYPE_INFO)
                End If
            ElseIf TipoConsulta = 2 Then
                If Not Ayuda.Existe(UserList(UserIndex).name) Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El mensaje ha sido entregado, ahora solo debes esperar que se desocupe algun GM." & FONTTYPE_INFO)
                    Call SendData(SendTarget.ToAdmins, 0, 0, "||Nuevo mensaje S.o.S" & "~195~250~35~1~0")
                    Call SendData(SendTarget.ToAdmins, 0, 0, "NEWSOSM[Denuncia / Acusacion]," & UserList(UserIndex).name & "," & rDatax)
                    Call Ayuda.Push(rData, UserList(UserIndex).name)
                Else
                    Call Ayuda.Quitar(UserList(UserIndex).name)
                    Call Ayuda.Push(rData, UserList(UserIndex).name)
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Ya habias mandado un mensaje, tu mensaje ha sido movido al final de la cola de mensajes." & FONTTYPE_INFO)
                End If
            ElseIf TipoConsulta = 3 Then
                If Not Ayuda.Existe(UserList(UserIndex).name) Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El mensaje ha sido entregado, ahora solo debes esperar que se desocupe algun GM." & FONTTYPE_INFO)
                    Call SendData(SendTarget.ToAdmins, 0, 0, "||Nuevo mensaje S.o.S" & "~195~250~35~1~0")
                    Call SendData(SendTarget.ToAdmins, 0, 0, "NEWSOSM[Sugerencia]," & UserList(UserIndex).name & "," & rDatax)
                    Call Ayuda.Push(rData, UserList(UserIndex).name)
                Else
                    Call Ayuda.Quitar(UserList(UserIndex).name)
                    Call Ayuda.Push(rData, UserList(UserIndex).name)
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Ya habias mandado un mensaje, tu mensaje ha sido movido al final de la cola de mensajes." & FONTTYPE_INFO)
                End If
            ElseIf TipoConsulta = 4 Then
                If Not Ayuda.Existe(UserList(UserIndex).name) Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El mensaje ha sido entregado, ahora solo debes esperar que se desocupe algun GM." & FONTTYPE_INFO)
                    Call SendData(SendTarget.ToAdmins, 0, 0, "NEWSOSM[Bug]," & UserList(UserIndex).name & "," & rDatax)
                    Call SendData(SendTarget.ToAdmins, 0, 0, "||Nuevo mensaje S.o.S" & "~195~250~35~1~0")
                    Call Ayuda.Push(rData, UserList(UserIndex).name)
                Else
                    Call Ayuda.Quitar(UserList(UserIndex).name)
                    Call Ayuda.Push(rData, UserList(UserIndex).name)
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Ya habias mandado un mensaje, tu mensaje ha sido movido al final de la cola de mensajes." & FONTTYPE_INFO)
                End If
            End If
            Exit Sub
 
 
        Case ";" 'Hablar
            rData = Right$(rData, Len(rData) - 1)
            If InStr(rData, "°") Then
                Exit Sub
            End If
        

'[Consejeros]
            
            If UserList(UserIndex).flags.Privilegios = PlayerType.Consejero Then
                Call LogGM(UserList(UserIndex).name, "Dijo: " & rData, True)
            End If
            
            ind = UserList(UserIndex).Char.CharIndex
            
            'piedra libre para todos los compas!
            If UserList(UserIndex).flags.Oculto > 0 Then
                UserList(UserIndex).flags.Oculto = 0
                If UserList(UserIndex).flags.Invisible = 0 Then
                    Call SendData(SendTarget.ToMap, 0, UserList(UserIndex).Pos.Map, "NOVER" & UserList(UserIndex).Char.CharIndex & ",0")
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡Has vuelto a ser visible!" & FONTTYPE_INFO)
                End If
            End If
            
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToDeadArea, UserIndex, UserList(UserIndex).Pos.Map, "||12632256°" & rData & "°" & CStr(ind))
            Else
                Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||" & vbWhite & "°" & rData & "°" & CStr(ind))
                'Call SendData(ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||" & UserList(UserIndex).name & "> " & rData & FONTTYPE_TALK)
            
            End If
        If UserList(UserIndex).flags.Privilegios = PlayerType.Admin Then
Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||" & vbYellow & "°" & rData & "°" & CStr(ind) & "°" & "1")
End If
            Exit Sub
        Case "-" 'Gritar
            If UserList(UserIndex).flags.Muerto = 1 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡Estas muerto!! Los muertos no pueden comunicarse con el mundo de los vivos. " & FONTTYPE_INFO)
                    Exit Sub
            End If
            rData = Right$(rData, Len(rData) - 1)
            If InStr(rData, "°") Then
                Exit Sub
            End If
            '[Consejeros]
            If UserList(UserIndex).flags.Privilegios = PlayerType.Consejero Then
                Call LogGM(UserList(UserIndex).name, "Grito: " & rData, True)
            End If
    
            'piedra libre para todos los compas!
            If UserList(UserIndex).flags.Oculto > 0 Then
                UserList(UserIndex).flags.Oculto = 0
                If UserList(UserIndex).flags.Invisible = 0 Then
                    Call SendData(SendTarget.ToMap, 0, UserList(UserIndex).Pos.Map, "NOVER" & UserList(UserIndex).Char.CharIndex & ",0")
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡Has vuelto a ser visible!" & FONTTYPE_INFO)
                End If
            End If
    
    
            ind = UserList(UserIndex).Char.CharIndex
            Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||" & vbRed & "°" & rData & "°" & str(ind))
            Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||" & UserList(UserIndex).name & "> " & rData & FONTTYPE_FIGHT)
            Exit Sub
            
              Case ":" 'Global
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡Estas muerto!! Los muertos no pueden comunicarse con el mundo de los vivos. " & FONTTYPE_INFO)
            Exit Sub
        End If
            rData = Right$(rData, Len(rData) - 1)
 
        If ChatGlobal = False Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El Global no esta activado" & FONTTYPE_INFO)
            Exit Sub
        End If
        If UserList(UserIndex).Stats.ELV < 45 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes mandar global si sos menor al nivel 45." & FONTTYPE_INFO)
        Exit Sub
        End If
        If UserList(UserIndex).Stats.GLD < 10000 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Para mandar un mensaje global necesitas 10.000 monedas de oro." & FONTTYPE_INFO)
        Exit Sub
        End If
        
        UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - 10000
        SendUserStatsBox (UserIndex)
        Call SendData(SendTarget.ToAll, 0, 0, "||[Global] " & UserList(UserIndex).name & "> " & rData & FONTTYPE_CELESTECN)
       
       
       Case "*" 'Mensaje privado
            rData = Right$(rData, Len(rData) - 1)
            tName = ReadField(1, rData, 32)
           
            'A los dioses y admins no vale susurrarles si no sos uno vos mismo (así no pueden ver si están conectados o no)
            If (EsDios(tName) Or EsAdmin(tName)) And UserList(UserIndex).flags.Privilegios < PlayerType.Dios Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes hablarle a los Dioses y Admins." & FONTTYPE_INFO)
                Exit Sub
            End If
           
            'A los Consejeros y SemiDioses no vale susurrarles si sos un PJ común.
            If UserList(UserIndex).flags.Privilegios = PlayerType.User And (EsSemiDios(tName) Or EsConsejero(tName)) Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes hablarle a los GMs" & FONTTYPE_INFO)
                Exit Sub
            End If
            If UserIndex = tIndex Then
               Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No podés chatear con vos mismo." & FONTTYPE_INFO)
               Exit Sub
            End If
            tIndex = NameIndex(tName)
            If tIndex <> 0 Then
                If Len(rData) <> Len(tName) Then
                    tMessage = Right$(rData, Len(rData) - (1 + Len(tName)))
                Else
                    tMessage = " "
                End If
                ind = UserList(UserIndex).Char.CharIndex
                If InStr(tMessage, "°") Then
                    Exit Sub
                End If
                '[Consejeros]
                If UserList(UserIndex).flags.Privilegios = PlayerType.Consejero Then
                    Call LogGM(UserList(UserIndex).name, "Le dijo a '" & UserList(tIndex).name & "' " & tMessage, True)
                End If

                    Call SendData(SendTarget.ToIndex, UserIndex, UserList(UserIndex).Pos.Map, "||" & "Le dijiste a " & UserList(tIndex).name & ": " & tMessage & "~128~255~255~0~0")
                    Call SendData(SendTarget.ToIndex, tIndex, UserList(UserIndex).Pos.Map, "||" & UserList(UserIndex).name & " te dice: " & tMessage & "~200~255~0~0~0")
                Exit Sub
            End If
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Usuario inexistente u offline. " & FONTTYPE_INFO)
            Exit Sub
               
               
       Case "\" 'Mensaje privado
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡Estas muerto!! Los muertos no pueden comunicarse con el mundo de los vivos. " & FONTTYPE_INFO)
                Exit Sub
            End If
        
            rData = Right$(rData, Len(rData) - 1)
            tName = ReadField(1, rData, 32)
           
            'A los dioses y admins no vale susurrarles si no sos uno vos mismo (así no pueden ver si están conectados o no)
            If (EsDios(tName) Or EsAdmin(tName)) And UserList(UserIndex).flags.Privilegios < PlayerType.Dios Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes susurrarle a los Dioses y Admins." & FONTTYPE_INFO)
                Exit Sub
            End If
           
            'A los Consejeros y SemiDioses no vale susurrarles si sos un PJ común.
            If UserList(UserIndex).flags.Privilegios = PlayerType.User And (EsSemiDios(tName) Or EsConsejero(tName)) Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes susurrarle a los GMs" & FONTTYPE_INFO)
                Exit Sub
            End If
           
            tIndex = NameIndex(tName)
            If tIndex <> 0 Then
                If Len(rData) <> Len(tName) Then
                    tMessage = Right$(rData, Len(rData) - (1 + Len(tName)))
                Else
                    tMessage = " "
                End If
                ind = UserList(UserIndex).Char.CharIndex
                If InStr(tMessage, "°") Then
                    Exit Sub
                End If
                      If UserList(tIndex).flags.DeseoRecibirMSJ = 0 Then
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El usuario tiene los mensajes privados bloqueados. " & FONTTYPE_INFO)
 Exit Sub
  End If
                '[Consejeros]
                If UserList(UserIndex).flags.Privilegios = PlayerType.Consejero Then
                    Call LogGM(UserList(UserIndex).name, "Le dijo a '" & UserList(tIndex).name & "' " & tMessage, True)
                End If
               
                    Call SendData(SendTarget.ToIndex, UserIndex, UserList(UserIndex).Pos.Map, "||" & "Le has dicho a " & UserList(tIndex).name & "> " & tMessage & FONTTYPE_ROJOC)
                    Call SendData(SendTarget.ToIndex, tIndex, UserList(UserIndex).Pos.Map, "||" & UserList(UserIndex).name & " te ha dicho> " & tMessage & FONTTYPE_ROJOC)
                Exit Sub
            End If
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Usuario inexistente. " & FONTTYPE_INFO)
            Exit Sub
        
        Case "M" 'Moverse
                          If CuentaUsuario > 0 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Espera que la cuenta llegue a 0." & FONTTYPE_INFO)
                Exit Sub
            End If
            Dim dummy As Long
            Dim TempTick As Long
            If UserList(UserIndex).flags.TimesWalk >= 30 Then
                TempTick = GetTickCount And &H7FFFFFFF
                dummy = (TempTick - UserList(UserIndex).flags.StartWalk)
                If dummy < 6050 Then
                    If TempTick - UserList(UserIndex).flags.CountSH > 90000 Then
                        UserList(UserIndex).flags.CountSH = 0
                    End If
                    If Not UserList(UserIndex).flags.CountSH = 0 Then
                        dummy = 126000 \ dummy
                        Call LogHackAttemp("Tramposo SH: " & UserList(UserIndex).name & " , " & dummy)
                        Call SendData(SendTarget.ToAdmins, 0, 0, "||Servidor> " & UserList(UserIndex).name & " ha sido echado por el servidor por posible uso de SH." & FONTTYPE_SERVER)
                        Call CloseSocket(UserIndex)
                        Exit Sub
                    Else
                        UserList(UserIndex).flags.CountSH = TempTick
                    End If
                End If
                UserList(UserIndex).flags.StartWalk = TempTick
                UserList(UserIndex).flags.TimesWalk = 0
            End If
            
            UserList(UserIndex).flags.TimesWalk = UserList(UserIndex).flags.TimesWalk + 1
            
            rData = Right$(rData, Len(rData) - 1)
            
            'salida parche
            If UserList(UserIndex).Counters.Saliendo Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||/SALIR cancelado." & FONTTYPE_WARNING)
                UserList(UserIndex).Counters.Saliendo = False
                UserList(UserIndex).Counters.Salir = 0
            End If
            
            If UserList(UserIndex).flags.Paralizado = 0 Then
                If Not UserList(UserIndex).flags.Descansar And Not UserList(UserIndex).flags.Meditando Then
                    Call MoveUserChar(UserIndex, val(rData))
                ElseIf UserList(UserIndex).flags.Descansar Then
                    UserList(UserIndex).flags.Descansar = False
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "DOK")
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has dejado de descansar." & FONTTYPE_INFO)
                    Call MoveUserChar(UserIndex, val(rData))
                    End If
                If UserList(UserIndex).flags.Meditando Then
                Exit Sub
                End If
                
            Else    'paralizado
                '[CDT 17-02-2004] (<- emmmmm ?????)
                If Not UserList(UserIndex).flags.UltimoMensaje = 1 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No podes moverte porque estas paralizado." & FONTTYPE_INFO)
                    UserList(UserIndex).flags.UltimoMensaje = 1
                End If
                '[/CDT]
                UserList(UserIndex).flags.CountSH = 0
            End If
            
            If UserList(UserIndex).flags.Oculto = 1 Then
                If UCase$(UserList(UserIndex).Clase) <> "LADRON" Then
                    UserList(UserIndex).flags.Oculto = 0
                    If UserList(UserIndex).flags.Invisible = 0 Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has vuelto a ser visible." & FONTTYPE_INFO)
                        Call SendData(SendTarget.ToMap, 0, UserList(UserIndex).Pos.Map, "NOVER" & UserList(UserIndex).Char.CharIndex & ",0")
                    End If
                End If
            End If
            
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call Empollando(UserIndex)
            Else
                UserList(UserIndex).flags.EstaEmpo = 0
                UserList(UserIndex).EmpoCont = 0
            End If
            Exit Sub
    End Select
    
    Select Case UCase$(rData)
   Case "/CAER"
 
    If Not UserList(UserIndex).flags.Privilegios > PlayerType.SemiDios Then Exit Sub
 
    With UserList(UserIndex)
    If MapInfo(.Pos.Map).SeCaenItems = 0 Then
    MapInfo(.Pos.Map).SeCaenItems = 1
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Atencion> Los items se caeran en el mapa." & FONTTYPE_INFO)
    Else
    MapInfo(.Pos.Map).SeCaenItems = 0
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Atencion> Los items no se caeran en el mapa." & FONTTYPE_INFO)
    End If
    End With
    Exit Sub
    Case "/REGRESAR"
If UserList(UserIndex).flags.Muerto = 1 Then
    Dim DeDonde As WorldPos
    Select Case UCase$(UserList(UserIndex).Hogar)
        Case "ULLATHORPE"
            DeDonde = Ullathorpe
          Case "NIX"
            DeDonde = Ullathorpe
    End Select
    Call WarpUserChar(UserIndex, DeDonde.Map, DeDonde.X, DeDonde.Y, True)
    Else
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡Solo los muertos pueden Volver!" & FONTTYPE_WARNING)
    End If
    Exit Sub
        Case "RPU" 'Pedido de actualizacion de la posicion
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "PU" & UserList(UserIndex).Pos.X & "," & UserList(UserIndex).Pos.Y)
            Exit Sub
        Case "AT"
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡No podes atacar a nadie porque estas muerto!!. " & FONTTYPE_INFO)
                Exit Sub
            End If
            If Not UserList(UserIndex).flags.ModoCombate Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No estas en modo de combate, presiona la tecla ""C"" para pasar al modo combate. " & FONTTYPE_INFO)
            Else
                If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                    If ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).proyectil = 1 Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No podés usar asi esta arma." & FONTTYPE_INFO)
                        Exit Sub
                    End If
                End If
                Call UsuarioAtaca(UserIndex)
                
                'piedra libre para todos los compas!
                If UserList(UserIndex).flags.Oculto > 0 And UserList(UserIndex).flags.AdminInvisible = 0 Then
                    UserList(UserIndex).flags.Oculto = 0
                    If UserList(UserIndex).flags.Invisible = 0 Then
                        Call SendData(SendTarget.ToMap, 0, UserList(UserIndex).Pos.Map, "NOVER" & UserList(UserIndex).Char.CharIndex & ",0")
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡Has vuelto a ser visible!" & FONTTYPE_INFO)
                    End If
                End If
                
            End If
            Exit Sub
        Case "AG"
            If UserList(UserIndex).flags.Muerto = 1 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡Estas muerto!! Los muertos no pueden tomar objetos. " & FONTTYPE_INFO)
                    Exit Sub
            End If
            '[Consejeros]
            If UserList(UserIndex).flags.Privilegios = PlayerType.Consejero And Not UserList(UserIndex).flags.EsRolesMaster Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes tomar ningun objeto. " & FONTTYPE_INFO)
                Exit Sub
            End If
            Call GetObj(UserIndex)
            Exit Sub
        Case "TAB" 'Entrar o salir modo combate
            If UserList(UserIndex).flags.ModoCombate Then
                SendData SendTarget.ToIndex, UserIndex, 0, "||Has salido del modo de combate. " & FONTTYPE_INFO
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "COGOFF")
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "COGON")
                SendData SendTarget.ToIndex, UserIndex, 0, "||Has pasado al modo de combate. " & FONTTYPE_INFO
            End If
            UserList(UserIndex).flags.ModoCombate = Not UserList(UserIndex).flags.ModoCombate
            Exit Sub
        Case "SEG" 'Activa / desactiva el seguro
            If UserList(UserIndex).flags.Seguro Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Escribe /SEG para quitar el seguro" & FONTTYPE_FIGHT)
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "SEGON")
                UserList(UserIndex).flags.Seguro = Not UserList(UserIndex).flags.Seguro
            End If
            Exit Sub
                Case "CCANJE"
        Dim Premios As Integer, SX As String
SX = "PRM" & UBound(PremiosList) & ","
 
For Premios = 1 To UBound(PremiosList)
    SX = SX & PremiosList(Premios).ObjName & ","
Next Premios
 
Call SendData(SendTarget.ToIndex, UserIndex, 0, SX & UserList(UserIndex).Stats.PuntosDeTorneo)
Call SendData(SendTarget.ToIndex, UserIndex, 0, "INF" & PremiosList(val(rData)).ObjRequiere & "," & PremiosList(val(rData)).ObjMaxAt & "," & PremiosList(val(rData)).ObjMinAt & "," & PremiosList(val(rData)).ObjMaxdef & "," & PremiosList(val(rData)).ObjMindef & "," & PremiosList(val(rData)).ObjMaxAtMag & "," & PremiosList(val(rData)).ObjMinAtMag & "," & PremiosList(val(rData)).ObjMaxDefMag & "," & PremiosList(val(rData)).ObjMinDefMag & "," & PremiosList(val(rData)).ObjDescripcion)
'sistema de premios [Dylan.-]
Exit Sub
        
        Case "ACTUALIZAR"
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "PU" & UserList(UserIndex).Pos.X & "," & UserList(UserIndex).Pos.Y)
            Exit Sub
        Case "GLINFO"
            tStr = SendGuildLeaderInfo(UserIndex)
            If tStr = vbNullString Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "GL" & SendGuildsList(UserIndex))
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "LEADERI" & tStr)
            End If
            Exit Sub
        Case "ATRI"
            Call EnviarAtrib(UserIndex)
            Exit Sub
        Case "ESKI"
            Call EnviarSkills(UserIndex)
            Exit Sub
        Case "FEST" 'Mini estadisticas :)
            Call EnviarMiniEstadisticas(UserIndex)
            Exit Sub
        '[Alejo]
        Case "FINCOM"
            'User sale del modo COMERCIO
            UserList(UserIndex).flags.Comerciando = False
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "FINCOMOK")
            Exit Sub
        Case "FINCOMUSU"
            'Sale modo comercio Usuario
            If UserList(UserIndex).ComUsu.DestUsu > 0 And _
                UserList(UserList(UserIndex).ComUsu.DestUsu).ComUsu.DestUsu = UserIndex Then
                Call SendData(SendTarget.ToIndex, UserList(UserIndex).ComUsu.DestUsu, 0, "||" & UserList(UserIndex).name & " ha dejado de comerciar con vos." & FONTTYPE_TALK)
                Call FinComerciarUsu(UserList(UserIndex).ComUsu.DestUsu)
            End If
            
            Call FinComerciarUsu(UserIndex)
            Exit Sub
        '[KEVIN]---------------------------------------
        '******************************************************
        Case "INIBOV"
            Call SendUserStatsBox(UserIndex)
            Call IniciarDeposito(UserIndex)
            Exit Sub
        Case "FINBAN"
            'User sale del modo BANCO
            UserList(UserIndex).flags.Comerciando = False
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "FINBANOK")
            Exit Sub
        '-------------------------------------------------------
        '[/KEVIN]**************************************
        Case "COMUSUOK"
            'Aceptar el cambio
            Call AceptarComercioUsu(UserIndex)
            Exit Sub
        Case "COMUSUNO"
            'Rechazar el cambio
            If UserList(UserIndex).ComUsu.DestUsu > 0 Then
                If UserList(UserList(UserIndex).ComUsu.DestUsu).flags.UserLogged Then
                    Call SendData(SendTarget.ToIndex, UserList(UserIndex).ComUsu.DestUsu, 0, "||" & UserList(UserIndex).name & " ha rechazado tu oferta." & FONTTYPE_TALK)
                    Call FinComerciarUsu(UserList(UserIndex).ComUsu.DestUsu)
                End If
            End If
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Has rechazado la oferta del otro usuario." & FONTTYPE_TALK)
            Call FinComerciarUsu(UserIndex)
            Exit Sub
        '[/Alejo]
    
    
    End Select
    
    
    
    Select Case UCase$(Left$(rData, 2))
    Case "ZI"
        rData = Right$(rData, Len(rData) - 2)
        Dim Bait(1 To 2) As Byte
        Bait(1) = val(ReadField(1, rData, 44))
        Bait(2) = val(ReadField(2, rData, 44))
        
        Select Case Bait(2)
            Case 0
                Bait(2) = Bait(1) - 1
            Case 1
                Bait(2) = Bait(1) + 1
            Case 2
                Bait(2) = Bait(1) - 5
            Case 3
                Bait(2) = Bait(1) + 5
        End Select
        
        If Bait(2) > 0 And Bait(2) <= MAX_INVENTORY_SLOTS Then Call AcomodarItems(UserIndex, Bait(1), Bait(2))
        
        Exit Sub
    '    Case "/Z"
    '        Dim Pos As WorldPos, Pos2 As WorldPos
    '        Dim O As Obj
    '
    '        For LoopC = 1 To 100
    '            Pos = UserList(UserIndex).Pos
    '            O.Amount = 1
    '            O.ObjIndex = iORO
    '            'Exit For
    '            Call TirarOro(100000, UserIndex)
    '            'Call Tilelibre(Pos, Pos2)
    '            'If Pos2.x = 0 Or Pos2.y = 0 Then Exit For
    '
    '            'Call MakeObj(SendTarget.ToMap, 0, UserList(UserIndex).Pos.Map, O, Pos2.Map, Pos2.x, Pos2.y)
    '        Next LoopC
    '
    '        Exit Sub
        Case "TI" 'Tirar item
                If UserList(UserIndex).flags.Navegando = 1 Or _
                   UserList(UserIndex).flags.Muerto = 1 Or _
                   (UserList(UserIndex).flags.Privilegios = PlayerType.Consejero And Not UserList(UserIndex).flags.EsRolesMaster) Then Exit Sub
                   '[Consejeros]
                
                rData = Right$(rData, Len(rData) - 2)
                Arg1 = ReadField(1, rData, 44)
                Arg2 = ReadField(2, rData, 44)
                If val(Arg1) = FLAGORO Then
                    
                    Call TirarOro(val(Arg2), UserIndex)
                    
                    Call SendUserStatsBox(UserIndex)
                    Exit Sub
                Else
                    If val(Arg1) <= MAX_INVENTORY_SLOTS And val(Arg1) > 0 Then
                        If UserList(UserIndex).Invent.Object(val(Arg1)).ObjIndex = 0 Then
                                Exit Sub
                        End If
                        Call DropObj(UserIndex, val(Arg1), val(Arg2), UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y)
                    Else
                        Exit Sub
                    End If
                End If
                Exit Sub
        Case "LH" ' Lanzar hechizo
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡Estas muerto!!." & FONTTYPE_INFO)
                Exit Sub
            End If
            rData = Right$(rData, Len(rData) - 2)
            UserList(UserIndex).flags.Hechizo = val(rData)
            Exit Sub
        Case "LC" 'Click izquierdo
            rData = Right$(rData, Len(rData) - 2)
            Arg1 = ReadField(1, rData, 44)
            Arg2 = ReadField(2, rData, 44)
            If Not Numeric(Arg1) Or Not Numeric(Arg2) Then Exit Sub
            X = CInt(Arg1)
            Y = CInt(Arg2)
            Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
            Exit Sub
        Case "RC" 'Click derecho
            rData = Right$(rData, Len(rData) - 2)
            Arg1 = ReadField(1, rData, 44)
            Arg2 = ReadField(2, rData, 44)
            If Not Numeric(Arg1) Or Not Numeric(Arg2) Then Exit Sub
            X = CInt(Arg1)
            Y = CInt(Arg2)
            Call Accion(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
            Exit Sub
        Case "UK"
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡Estas muerto!!." & FONTTYPE_INFO)
                Exit Sub
            End If
    
            rData = Right$(rData, Len(rData) - 2)
            Select Case val(rData)
                Case Robar
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "T01" & Robar)
                Case Magia
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "T01" & Magia)
                Case Domar
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "T01" & Domar)
                Case Ocultarse
                    If UserList(UserIndex).flags.Navegando = 1 Then
                        '[CDT 17-02-2004]
                        If Not UserList(UserIndex).flags.UltimoMensaje = 3 Then
                            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No podes ocultarte si estas navegando." & FONTTYPE_INFO)
                            UserList(UserIndex).flags.UltimoMensaje = 3
                        End If
                        '[/CDT]
                        Exit Sub
                    End If
                    
                    If UserList(UserIndex).flags.Oculto = 1 Then
                        '[CDT 17-02-2004]
                        If Not UserList(UserIndex).flags.UltimoMensaje = 2 Then
                            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Ya estas oculto." & FONTTYPE_INFO)
                            UserList(UserIndex).flags.UltimoMensaje = 2
                        End If
                        '[/CDT]
                        Exit Sub
                    End If
                    
                    Call DoOcultarse(UserIndex)
            End Select
            Exit Sub
    
    End Select
    
    Select Case UCase$(Left$(rData, 3))
        Case "USA"
            rData = Right$(rData, Len(rData) - 3)
            If val(rData) <= MAX_INVENTORY_SLOTS And val(rData) > 0 Then
                If UserList(UserIndex).Invent.Object(val(rData)).ObjIndex = 0 Then Exit Sub
            Else
                Exit Sub
            End If
            If UserList(UserIndex).flags.Meditando Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "M!")
                Exit Sub
            End If
            Call UseInvItem(UserIndex, val(rData))
            Exit Sub
        Case "CNS" ' Construye herreria
            rData = Right$(rData, Len(rData) - 3)
            X = CInt(rData)
            If X < 1 Then Exit Sub
            If ObjData(X).SkHerreria = 0 Then Exit Sub
            Call HerreroConstruirItem(UserIndex, X)
            Exit Sub
        Case "CNC" ' Construye carpinteria
            rData = Right$(rData, Len(rData) - 3)
            X = CInt(rData)
            If X < 1 Or ObjData(X).SkCarpinteria = 0 Then Exit Sub
            Call CarpinteroConstruirItem(UserIndex, X)
            Exit Sub
        Case "WLC" 'Click izquierdo en modo trabajo
            rData = Right$(rData, Len(rData) - 3)
            Arg1 = ReadField(1, rData, 44)
            Arg2 = ReadField(2, rData, 44)
            Arg3 = ReadField(3, rData, 44)
            If Arg3 = "" Or Arg2 = "" Or Arg1 = "" Then Exit Sub
            If Not Numeric(Arg1) Or Not Numeric(Arg2) Or Not Numeric(Arg3) Then Exit Sub
            
            X = CInt(Arg1)
            Y = CInt(Arg2)
            tLong = CInt(Arg3)
            
            If UserList(UserIndex).flags.Muerto = 1 Or _
               UserList(UserIndex).flags.Descansar Or _
               UserList(UserIndex).flags.Meditando Or _
               Not InMapBounds(X, Y) Then Exit Sub
            
            If Not InRangoVision(UserIndex, X, Y) Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "PU" & UserList(UserIndex).Pos.X & "," & UserList(UserIndex).Pos.Y)
                Exit Sub
            End If
            
            Select Case tLong
            
            Case Proyectiles
                Dim TU As Integer, tN As Integer
                'Nos aseguramos que este usando un arma de proyectiles
                If Not IntervaloPermiteAtacar(UserIndex, False) Or Not IntervaloPermiteUsarArcos(UserIndex) Then
                    Exit Sub
                End If

                DummyInt = 0

                If UserList(UserIndex).Invent.WeaponEqpObjIndex = 0 Then
                    DummyInt = 1
                ElseIf UserList(UserIndex).Invent.WeaponEqpSlot < 1 Or UserList(UserIndex).Invent.WeaponEqpSlot > MAX_INVENTORY_SLOTS Then
                    DummyInt = 1
                ElseIf UserList(UserIndex).Invent.MunicionEqpSlot < 1 Or UserList(UserIndex).Invent.MunicionEqpSlot > MAX_INVENTORY_SLOTS Then
                    DummyInt = 1
                ElseIf UserList(UserIndex).Invent.MunicionEqpObjIndex = 0 Then
                    DummyInt = 1
                ElseIf ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).proyectil <> 1 Then
                    DummyInt = 2
                ElseIf ObjData(UserList(UserIndex).Invent.MunicionEqpObjIndex).OBJType <> eOBJType.otFLECHAS Then
                    DummyInt = 1
                ElseIf UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.MunicionEqpSlot).Amount < 1 Then
                    DummyInt = 1
                End If
                
                If DummyInt <> 0 Then
                    If DummyInt = 1 Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No tenes municiones." & FONTTYPE_INFO)
                    End If
                    Call Desequipar(UserIndex, UserList(UserIndex).Invent.MunicionEqpSlot)
                    Call Desequipar(UserIndex, UserList(UserIndex).Invent.WeaponEqpSlot)
                    Exit Sub
                End If
                
                DummyInt = 0
                'Quitamos stamina
                If UserList(UserIndex).Stats.MinSta >= 10 Then
                     Call QuitarSta(UserIndex, RandomNumber(1, 10))
                Else
                     Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas muy cansado para luchar." & FONTTYPE_INFO)
                     Exit Sub
                End If
                 
                Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, Arg1, Arg2)
                
                TU = UserList(UserIndex).flags.TargetUser
                tN = UserList(UserIndex).flags.TargetNPC
                
                'Sólo permitimos atacar si el otro nos puede atacar también
                If TU > 0 Then
                    If Abs(UserList(UserList(UserIndex).flags.TargetUser).Pos.Y - UserList(UserIndex).Pos.Y) > RANGO_VISION_Y Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas demasiado lejos para atacar." & FONTTYPE_WARNING)
                        Exit Sub
                    End If
                ElseIf tN > 0 Then
                    If Abs(Npclist(UserList(UserIndex).flags.TargetNPC).Pos.Y - UserList(UserIndex).Pos.Y) > RANGO_VISION_Y Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas demasiado lejos para atacar." & FONTTYPE_WARNING)
                        Exit Sub
                    End If
                End If
                
                If UserList(UserIndex).flags.Privilegios <= PlayerType.Consejero Then
                If TU > 0 Then
                    'Previene pegarse a uno mismo
                    If TU = UserIndex Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡No puedes atacarte a vos mismo!" & FONTTYPE_INFO)
                        DummyInt = 1
                        Exit Sub
                    End If
                End If
    End If
    
                If DummyInt = 0 Then
                    'Saca 1 flecha
                    DummyInt = UserList(UserIndex).Invent.MunicionEqpSlot
                    Call QuitarUserInvItem(UserIndex, UserList(UserIndex).Invent.MunicionEqpSlot, 1)
                    If DummyInt < 1 Or DummyInt > MAX_INVENTORY_SLOTS Then Exit Sub
                    If UserList(UserIndex).Invent.Object(DummyInt).Amount > 0 Then
                        UserList(UserIndex).Invent.Object(DummyInt).Equipped = 1
                        UserList(UserIndex).Invent.MunicionEqpSlot = DummyInt
                        UserList(UserIndex).Invent.MunicionEqpObjIndex = UserList(UserIndex).Invent.Object(DummyInt).ObjIndex
                        Call UpdateUserInv(False, UserIndex, UserList(UserIndex).Invent.MunicionEqpSlot)
                    Else
                        Call UpdateUserInv(False, UserIndex, DummyInt)
                        UserList(UserIndex).Invent.MunicionEqpSlot = 0
                        UserList(UserIndex).Invent.MunicionEqpObjIndex = 0
                    End If
                    '-----------------------------------
                End If

                If tN > 0 Then
                    If Npclist(tN).Attackable <> 0 Then
                        Call UsuarioAtacaNpc(UserIndex, tN)
                    End If
                ElseIf TU > 0 Then
                    If UserList(UserIndex).flags.Seguro Then
                        If Not Criminal(TU) Then
                            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡Para atacar ciudadanos desactiva el seguro!" & FONTTYPE_FIGHT)
                            Exit Sub
                        End If
                    End If
                    Call UsuarioAtacaUsuario(UserIndex, TU)
                End If
                
            Case Magia
                If MapInfo(UserList(UserIndex).Pos.Map).MagiaSinEfecto > 0 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Una fuerza oscura te impide canalizar tu energía" & FONTTYPE_FIGHT)
                    Exit Sub
                End If
                Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
                
                'MmMmMmmmmM
                Dim wp2 As WorldPos
                wp2.Map = UserList(UserIndex).Pos.Map
                wp2.X = X
                wp2.Y = Y
                                
                If UserList(UserIndex).flags.Hechizo > 0 Then
                    If IntervaloPermiteLanzarSpell(UserIndex) Then
                        Call LanzarHechizo(UserList(UserIndex).flags.Hechizo, UserIndex)
                    '    UserList(UserIndex).flags.PuedeLanzarSpell = 0
                        UserList(UserIndex).flags.Hechizo = 0
                    End If
                Else
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡Primero selecciona el hechizo que quieres lanzar!" & FONTTYPE_INFO)
                End If
                
                'If Distancia(UserList(UserIndex).Pos, wp2) > 10 Then
                If (Abs(UserList(UserIndex).Pos.X - wp2.X) > 9 Or Abs(UserList(UserIndex).Pos.Y - wp2.Y) > 8) Then
                    Dim txt As String
                    txt = "Ataque fuera de rango de " & UserList(UserIndex).name & "(" & UserList(UserIndex).Pos.Map & "/" & UserList(UserIndex).Pos.X & "/" & UserList(UserIndex).Pos.Y & ") ip: " & UserList(UserIndex).ip & " a la posicion (" & wp2.Map & "/" & wp2.X & "/" & wp2.Y & ") "
                    If UserList(UserIndex).flags.Hechizo > 0 Then
                        txt = txt & ". Hechizo: " & Hechizos(UserList(UserIndex).Stats.UserHechizos(UserList(UserIndex).flags.Hechizo)).Nombre
                    End If
                    If MapData(wp2.Map, wp2.X, wp2.Y).UserIndex > 0 Then
                        txt = txt & " hacia el usuario: " & UserList(MapData(wp2.Map, wp2.X, wp2.Y).UserIndex).name
                    ElseIf MapData(wp2.Map, wp2.X, wp2.Y).NpcIndex > 0 Then
                        txt = txt & " hacia el NPC: " & Npclist(MapData(wp2.Map, wp2.X, wp2.Y).NpcIndex).name
                    End If
                    
                    Call LogCheating(txt)
                End If
                
            
            
            
            Case Pesca
                        
                AuxInd = UserList(UserIndex).Invent.HerramientaEqpObjIndex
                If AuxInd = 0 Then Exit Sub
                
                'If UserList(UserIndex).flags.PuedeTrabajar = 0 Then Exit Sub
                If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub
                
                If AuxInd <> CAÑA_PESCA And AuxInd <> RED_PESCA Then
                    'Call Cerrar_Usuario(UserIndex)
                    ' Podemos llegar acá si el user equipó el anillo dsp de la U y antes del click
                    Exit Sub
                End If
                
                'Basado en la idea de Barrin
                'Comentario por Barrin: jah, "basado", caradura ! ^^
                If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = 1 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes pescar desde donde te encuentras." & FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If HayAgua(UserList(UserIndex).Pos.Map, X, Y) Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_PESCAR)
                    
                    Select Case AuxInd
                    Case CAÑA_PESCA
                        Call DoPescar(UserIndex)
                    Case RED_PESCA
                        With UserList(UserIndex)
                            wpaux.Map = .Pos.Map
                            wpaux.X = X
                            wpaux.Y = Y
                        End With
                        
                        If Distancia(UserList(UserIndex).Pos, wpaux) > 2 Then
                            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estás demasiado lejos para pescar." & FONTTYPE_INFO)
                            Exit Sub
                        End If
                        
                        Call DoPescarRed(UserIndex)
                    End Select
    
                Else
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No hay agua donde pescar busca un lago, rio o mar." & FONTTYPE_INFO)
                End If
                
            Case Robar
               If MapInfo(UserList(UserIndex).Pos.Map).Pk Then
                    'If UserList(UserIndex).flags.PuedeTrabajar = 0 Then Exit Sub
                    If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub
                    
                    Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
                    
                    If UserList(UserIndex).flags.TargetUser > 0 And UserList(UserIndex).flags.TargetUser <> UserIndex Then
                       If UserList(UserList(UserIndex).flags.TargetUser).flags.Muerto = 0 Then
                            wpaux.Map = UserList(UserIndex).Pos.Map
                            wpaux.X = val(ReadField(1, rData, 44))
                            wpaux.Y = val(ReadField(2, rData, 44))
                            If Distancia(wpaux, UserList(UserIndex).Pos) > 2 Then
                                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                                Exit Sub
                            End If
                            '17/09/02
                            'No aseguramos que el trigger le permite robar
                            If MapData(UserList(UserList(UserIndex).flags.TargetUser).Pos.Map, UserList(UserList(UserIndex).flags.TargetUser).Pos.X, UserList(UserList(UserIndex).flags.TargetUser).Pos.Y).trigger = eTrigger.ZONASEGURA Then
                                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No podes robar aquí." & FONTTYPE_WARNING)
                                Exit Sub
                            End If
                            If MapData(UserList(UserIndex).Pos.Map, UserList(UserIndex).Pos.X, UserList(UserIndex).Pos.Y).trigger = eTrigger.ZONASEGURA Then
                                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No podes robar aquí." & FONTTYPE_WARNING)
                                Exit Sub
                            End If
                            
                            Call DoRobar(UserIndex, UserList(UserIndex).flags.TargetUser)
                       End If
                    Else
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No a quien robarle!." & FONTTYPE_INFO)
                    End If
                Else
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡No podes robarle en zonas seguras!." & FONTTYPE_INFO)
                End If
            Case Talar
                
                'If UserList(UserIndex).flags.PuedeTrabajar = 0 Then Exit Sub
                If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub
                
                If UserList(UserIndex).Invent.HerramientaEqpObjIndex = 0 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Deberías equiparte el hacha." & FONTTYPE_INFO)
                    Exit Sub
                End If
                
                If UserList(UserIndex).Invent.HerramientaEqpObjIndex <> HACHA_LEÑADOR Then
                    ' Call Cerrar_Usuario(UserIndex)
                    ' Podemos llegar acá si el user equipó el anillo dsp de la U y antes del click
                    Exit Sub
                End If
                
                AuxInd = MapData(UserList(UserIndex).Pos.Map, X, Y).OBJInfo.ObjIndex
                If AuxInd > 0 Then
                    wpaux.Map = UserList(UserIndex).Pos.Map
                    wpaux.X = X
                    wpaux.Y = Y
                    If Distancia(wpaux, UserList(UserIndex).Pos) > 2 Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    'Barrin 29/9/03
                    If Distancia(wpaux, UserList(UserIndex).Pos) = 0 Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No podes talar desde allí." & FONTTYPE_INFO)
                        Exit Sub
                    End If
                    
                    '¿Hay un arbol donde clickeo?
                    If ObjData(AuxInd).OBJType = eOBJType.otArboles Then
                        Call SendData(SendTarget.ToPCArea, CInt(UserIndex), UserList(UserIndex).Pos.Map, "TW" & SND_TALAR)
                        Call DoTalar(UserIndex)
                    End If
                Else
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No hay ningun arbol ahi." & FONTTYPE_INFO)
                End If
            Case Mineria
                
                'If UserList(UserIndex).flags.PuedeTrabajar = 0 Then Exit Sub
                If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub
                                
                If UserList(UserIndex).Invent.HerramientaEqpObjIndex = 0 Then Exit Sub
                
                If UserList(UserIndex).Invent.HerramientaEqpObjIndex <> PIQUETE_MINERO Then
                    ' Call Cerrar_Usuario(UserIndex)
                    ' Podemos llegar acá si el user equipó el anillo dsp de la U y antes del click
                    Exit Sub
                End If
                
                Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
                
                AuxInd = MapData(UserList(UserIndex).Pos.Map, X, Y).OBJInfo.ObjIndex
                If AuxInd > 0 Then
                    wpaux.Map = UserList(UserIndex).Pos.Map
                    wpaux.X = X
                    wpaux.Y = Y
                    If Distancia(wpaux, UserList(UserIndex).Pos) > 2 Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                        Exit Sub
                    End If
                    '¿Hay un yacimiento donde clickeo?
                    If ObjData(AuxInd).OBJType = eOBJType.otYacimiento Then
                        Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "TW" & SND_MINERO)
                        Call DoMineria(UserIndex)
                    Else
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Ahi no hay ningun yacimiento." & FONTTYPE_INFO)
                    End If
                Else
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Ahi no hay ningun yacimiento." & FONTTYPE_INFO)
                End If
            Case Domar
              'Modificado 25/11/02
              'Optimizado y solucionado el bug de la doma de
              'criaturas hostiles.
              Dim CI As Integer
              
              Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
              CI = UserList(UserIndex).flags.TargetNPC
              
              If CI > 0 Then
                       If Npclist(CI).flags.Domable > 0 Then
                            wpaux.Map = UserList(UserIndex).Pos.Map
                            wpaux.X = X
                            wpaux.Y = Y
                            If Distancia(wpaux, Npclist(UserList(UserIndex).flags.TargetNPC).Pos) > 2 Then
                                  Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Estas demasiado lejos." & FONTTYPE_INFO)
                                  Exit Sub
                            End If
                            If Npclist(CI).flags.AttackedBy <> "" Then
                                  Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No podés domar una criatura que está luchando con un jugador." & FONTTYPE_INFO)
                                  Exit Sub
                            End If
                            Call DoDomar(UserIndex, CI)
                        Else
                            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No podes domar a esa criatura." & FONTTYPE_INFO)
                        End If
              Else
                     Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No hay ninguna criatura alli!." & FONTTYPE_INFO)
              End If
              
            Case FundirMetal
                'Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
                If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub
                
                If UserList(UserIndex).flags.TargetObj > 0 Then
                    If ObjData(UserList(UserIndex).flags.TargetObj).OBJType = eOBJType.otFragua Then
                        ''chequeamos que no se zarpe duplicando oro
                        If UserList(UserIndex).Invent.Object(UserList(UserIndex).flags.TargetObjInvSlot).ObjIndex <> UserList(UserIndex).flags.TargetObjInvIndex Then
                            If UserList(UserIndex).Invent.Object(UserList(UserIndex).flags.TargetObjInvSlot).ObjIndex = 0 Or UserList(UserIndex).Invent.Object(UserList(UserIndex).flags.TargetObjInvSlot).Amount = 0 Then
                                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No tienes mas minerales" & FONTTYPE_INFO)
                                Exit Sub
                            End If
                            
                            ''FUISTE
                            'Call Ban(UserList(UserIndex).Name, "Sistema anti cheats", "Intento de duplicacion de items")
                            'Call LogCheating(UserList(UserIndex).Name & " intento crear minerales a partir de otros: FlagSlot/usaba/usoconclick/cantidad/IP:" & UserList(UserIndex).flags.TargetObjInvSlot & "/" & UserList(UserIndex).flags.TargetObjInvIndex & "/" & UserList(UserIndex).Invent.Object(UserList(UserIndex).flags.TargetObjInvSlot).ObjIndex & "/" & UserList(UserIndex).Invent.Object(UserList(UserIndex).flags.TargetObjInvSlot).Amount & "/" & UserList(UserIndex).ip)
                            'UserList(UserIndex).flags.Ban = 1
                            'Call SendData(SendTarget.ToAll, 0, 0, "||>>>> El sistema anti-cheats baneó a " & UserList(UserIndex).Name & " (intento de duplicación). Ip Logged. " & FONTTYPE_FIGHT)
                            Call SendData(SendTarget.ToIndex, UserIndex, 0, "ERRHas sido expulsado por el sistema anti cheats. Reconéctate.")
                            Call CloseSocket(UserIndex)
                            Exit Sub
                        End If
                        Call FundirMineral(UserIndex)
                    Else
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Ahi no hay ninguna fragua." & FONTTYPE_INFO)
                    End If
                Else
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Ahi no hay ninguna fragua." & FONTTYPE_INFO)
                End If
                
            Case Herreria
                Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
                
                If UserList(UserIndex).flags.TargetObj > 0 Then
                    If ObjData(UserList(UserIndex).flags.TargetObj).OBJType = eOBJType.otYunque Then
                        Call EnivarArmasConstruibles(UserIndex)
                        Call EnivarArmadurasConstruibles(UserIndex)
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "SFH")
                    Else
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Ahi no hay ningun yunque." & FONTTYPE_INFO)
                    End If
                Else
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Ahi no hay ningun yunque." & FONTTYPE_INFO)
                End If
                
            End Select
            
            'UserList(UserIndex).flags.PuedeTrabajar = 0
            Exit Sub
        Case "CIG"
            rData = Right$(rData, Len(rData) - 3)
            
            If modGuilds.CrearNuevoClan(rData, UserIndex, UserList(UserIndex).FundandoGuildAlineacion, tStr) Then
            Call QuitarObjetos(875, 1, UserIndex)
                Call SendData(SendTarget.ToAll, 0, 0, "||" & UserList(UserIndex).name & " fundó el clan " & Guilds(UserList(UserIndex).GuildIndex).GuildName & " de alineación " & Alineacion2String(Guilds(UserList(UserIndex).GuildIndex).Alineacion) & "." & FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & tStr & FONTTYPE_GUILD)
            End If
            
            Exit Sub
    End Select
    
    
    
    
    
    Select Case UCase$(Left$(rData, 4))
               Case "VPSR"
                Dim Proceso As String
                rData = Right$(rData, Len(rData) - 4)
                Proceso = ReadField(1, rData, 44)
                tIndex = ReadField(2, rData, 44)
                Call SendData(SendTarget.ToIndex, tIndex, 0, "PCSE" & Proceso)
            Exit Sub
            Case "TALX"
                Dim ProMata As String
                rData = Right$(rData, Len(rData) - 4)
                ProMata = ReadField(1, rData, 44)
                tIndex = ReadField(2, rData, 44)
                Call SendData(SendTarget.ToIndex, tIndex, 0, "GFSE" & ProMata)
            Exit Sub
    Case "PCCC" 'Te veo el caption jaja esa eM
            Dim caption As String
            rData = Right$(rData, Len(rData) - 4)
            caption = ReadField(1, rData, 44)
            tIndex = ReadField(2, rData, 44)
            Call SendData(SendTarget.ToIndex, tIndex, 0, "PCCC" & caption & "," & UserList(UserIndex).name)
            Exit Sub
           Case "SWAP" ' Te muevo el item
            rData = Right$(rData, Len(rData) - 4)
            ObjSlot1 = ReadField(1, rData, 44)
            ObjSlot2 = ReadField(2, rData, 44)
            SwapObjects (UserIndex)
            Exit Sub
        Case "INFS" 'Informacion del hechizo
                rData = Right$(rData, Len(rData) - 4)
                If val(rData) > 0 And val(rData) < MAXUSERHECHIZOS + 1 Then
                    Dim H As Integer
                    H = UserList(UserIndex).Stats.UserHechizos(val(rData))
                    If H > 0 And H < NumeroHechizos + 1 Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||%%%%%%%%%%%% INFO DEL HECHIZO %%%%%%%%%%%%" & FONTTYPE_INFO)
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Nombre:" & Hechizos(H).Nombre & FONTTYPE_INFO)
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Descripcion:" & Hechizos(H).Desc & FONTTYPE_INFO)
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Skill requerido: " & Hechizos(H).MinSkill & " de magia." & FONTTYPE_INFO)
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Mana necesario: " & Hechizos(H).ManaRequerido & FONTTYPE_INFO)
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Stamina necesaria: " & Hechizos(H).StaRequerido & FONTTYPE_INFO)
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%" & FONTTYPE_INFO)
                    End If
                Else
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡Primero selecciona el hechizo.!" & FONTTYPE_INFO)
                End If
                Exit Sub
        Case "EQUI"
                If UserList(UserIndex).flags.Muerto = 1 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡Estas muerto!! Solo podes usar items cuando estas vivo. " & FONTTYPE_INFO)
                    Exit Sub
                End If
                rData = Right$(rData, Len(rData) - 4)
                If val(rData) <= MAX_INVENTORY_SLOTS And val(rData) > 0 Then
                     If UserList(UserIndex).Invent.Object(val(rData)).ObjIndex = 0 Then Exit Sub
                Else
                    Exit Sub
                End If
                Call EquiparInvItem(UserIndex, val(rData))
                Exit Sub
        Case "CHEA" 'Cambiar Heading ;-)
            rData = Right$(rData, Len(rData) - 4)
            If val(rData) > 0 And val(rData) < 5 Then
                UserList(UserIndex).Char.Heading = rData
                Call ChangeUserChar(SendTarget.ToMap, 0, UserList(UserIndex).Pos.Map, UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
            End If
            Exit Sub
        Case "SKSE" 'Modificar skills
            Dim sumatoria As Integer
            Dim incremento As Integer
            rData = Right$(rData, Len(rData) - 4)
            
            'Codigo para prevenir el hackeo de los skills
            '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
            For i = 1 To NUMSKILLS
                incremento = val(ReadField(i, rData, 44))
                
                If incremento < 0 Then
                    'Call SendData(SendTarget.ToAll, 0, 0, "||Los Dioses han desterrado a " & UserList(UserIndex).Name & FONTTYPE_INFO)
                    Call LogHackAttemp(UserList(UserIndex).name & " IP:" & UserList(UserIndex).ip & " trato de hackear los skills.")
                    UserList(UserIndex).Stats.SkillPts = 0
                    Call CloseSocket(UserIndex)
                    Exit Sub
                End If
                
                sumatoria = sumatoria + incremento
            Next i
            
            If sumatoria > UserList(UserIndex).Stats.SkillPts Then
                'UserList(UserIndex).Flags.AdministrativeBan = 1
                'Call SendData(SendTarget.ToAll, 0, 0, "||Los Dioses han desterrado a " & UserList(UserIndex).Name & FONTTYPE_INFO)
                Call LogHackAttemp(UserList(UserIndex).name & " IP:" & UserList(UserIndex).ip & " trato de hackear los skills.")
                Call CloseSocket(UserIndex)
                Exit Sub
            End If
            '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
            
            For i = 1 To NUMSKILLS
                incremento = val(ReadField(i, rData, 44))
                UserList(UserIndex).Stats.SkillPts = UserList(UserIndex).Stats.SkillPts - incremento
                UserList(UserIndex).Stats.UserSkills(i) = UserList(UserIndex).Stats.UserSkills(i) + incremento
                If UserList(UserIndex).Stats.UserSkills(i) > 100 Then UserList(UserIndex).Stats.UserSkills(i) = 100
            Next i
            Exit Sub
        Case "ENTR" 'Entrena hombre!
            
            If UserList(UserIndex).flags.TargetNPC = 0 Then Exit Sub
            
            If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> 3 Then Exit Sub
            
            rData = Right$(rData, Len(rData) - 4)
            
            If Npclist(UserList(UserIndex).flags.TargetNPC).Mascotas < MAXMASCOTASENTRENADOR Then
                If val(rData) > 0 And val(rData) < Npclist(UserList(UserIndex).flags.TargetNPC).NroCriaturas + 1 Then
                        Dim SpawnedNpc As Integer
                        SpawnedNpc = SpawnNpc(Npclist(UserList(UserIndex).flags.TargetNPC).Criaturas(val(rData)).NpcIndex, Npclist(UserList(UserIndex).flags.TargetNPC).Pos, True, False)
                        If SpawnedNpc > 0 Then
                            Npclist(SpawnedNpc).MaestroNpc = UserList(UserIndex).flags.TargetNPC
                            Npclist(UserList(UserIndex).flags.TargetNPC).Mascotas = Npclist(UserList(UserIndex).flags.TargetNPC).Mascotas + 1
                        End If
                End If
            Else
                Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||" & vbWhite & "°" & "No puedo traer mas criaturas, mata las existentes!" & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
            End If
            
            Exit Sub

        Case "COMP"
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                Exit Sub
            End If
            
            '¿El target es un NPC valido?
            If UserList(UserIndex).flags.TargetNPC > 0 Then
                '¿El NPC puede comerciar?
                If Npclist(UserList(UserIndex).flags.TargetNPC).Comercia = 0 Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||" & FONTTYPE_TALK & "°" & "No tengo ningun interes en comerciar." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
                    Exit Sub
                End If
            Else
                Exit Sub
            End If
            rData = Right$(rData, Len(rData) - 5)
            'User compra el item del slot rdata
            If UserList(UserIndex).flags.Comerciando = False Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No estas comerciando " & FONTTYPE_INFO)
                Exit Sub
            End If
            'listindex+1, cantidad
            Call NPCVentaItem(UserIndex, val(ReadField(1, rData, 44)), val(ReadField(2, rData, 44)), UserList(UserIndex).flags.TargetNPC)
            Exit Sub
        '[KEVIN]*********************************************************************
        '------------------------------------------------------------------------------------
        Case "RETI"
             '¿Esta el user muerto? Si es asi no puede comerciar
             If UserList(UserIndex).flags.Muerto = 1 Then
                       Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                       Exit Sub
             End If
             '¿El target es un NPC valido?
             If UserList(UserIndex).flags.TargetNPC > 0 Then
                   '¿Es el banquero?
                   If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> 4 Then
                       Exit Sub
                   End If
             Else
               Exit Sub
             End If
             rData = Right(rData, Len(rData) - 5)
             'User retira el item del slot rdata
             Call UserRetiraItem(UserIndex, val(ReadField(1, rData, 44)), val(ReadField(2, rData, 44)))
             Exit Sub
        '-----------------------------------------------------------------------------------
        '[/KEVIN]****************************************************************************
        Case "VEND"
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                Exit Sub
            End If
            rData = Right$(rData, Len(rData) - 5)
            '¿El target es un NPC valido?
            tInt = val(ReadField(1, rData, 44))
            If UserList(UserIndex).flags.TargetNPC > 0 Then
                '¿El NPC puede comerciar?
                If Npclist(UserList(UserIndex).flags.TargetNPC).Comercia = 0 Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, UserList(UserIndex).Pos.Map, "||" & FONTTYPE_TALK & "°" & "No tengo ningun interes en comerciar." & "°" & str(Npclist(UserList(UserIndex).flags.TargetNPC).Char.CharIndex))
                    Exit Sub
                End If
            Else
                Exit Sub
            End If
'           rdata = Right$(rdata, Len(rdata) - 5)
            'User compra el item del slot rdata
            Call NPCCompraItem(UserIndex, val(ReadField(1, rData, 44)), val(ReadField(2, rData, 44)))
            Exit Sub
        '[KEVIN]-------------------------------------------------------------------------
        '****************************************************************************************
        Case "DEPO"
            '¿Esta el user muerto? Si es asi no puede comerciar
            If UserList(UserIndex).flags.Muerto = 1 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||¡¡Estas muerto!!" & FONTTYPE_INFO)
                Exit Sub
            End If
            '¿El target es un NPC valido?
            If UserList(UserIndex).flags.TargetNPC > 0 Then
                '¿El NPC puede comerciar?
                If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> eNPCType.Banquero Then
                    Exit Sub
                End If
            Else
                Exit Sub
            End If
            rData = Right(rData, Len(rData) - 5)
            'User deposita el item del slot rdata
            Call UserDepositaItem(UserIndex, val(ReadField(1, rData, 44)), val(ReadField(2, rData, 44)))
            Exit Sub
        '****************************************************************************************
        '[/KEVIN]---------------------------------------------------------------------------------
    End Select
    Select Case UCase$(Left$(rData, 5))
        Case "DEMSG"
            If UserList(UserIndex).flags.TargetObj > 0 Then
            rData = Right$(rData, Len(rData) - 5)
            Dim f As String, Titu As String, msg As String, f2 As String
            f = App.Path & "\foros\"
            f = f & UCase$(ObjData(UserList(UserIndex).flags.TargetObj).ForoID) & ".for"
            Titu = ReadField(1, rData, 176)
            msg = ReadField(2, rData, 176)
            Dim n2 As Integer, loopme As Integer
            If FileExist(f, vbNormal) Then
                Dim num As Integer
                num = val(GetVar(f, "INFO", "CantMSG"))
                If num > MAX_MENSAJES_FORO Then
                    For loopme = 1 To num
                        Kill App.Path & "\foros\" & UCase$(ObjData(UserList(UserIndex).flags.TargetObj).ForoID) & loopme & ".for"
                    Next
                    Kill App.Path & "\foros\" & UCase$(ObjData(UserList(UserIndex).flags.TargetObj).ForoID) & ".for"
                    num = 0
                End If
                n2 = FreeFile
                f2 = Left$(f, Len(f) - 4)
                f2 = f2 & num + 1 & ".for"
                Open f2 For Output As n2
                Print #n2, Titu
                Print #n2, msg
                Call WriteVar(f, "INFO", "CantMSG", num + 1)
            Else
                n2 = FreeFile
                f2 = Left$(f, Len(f) - 4)
                f2 = f2 & "1" & ".for"
                Open f2 For Output As n2
                Print #n2, Titu
                Print #n2, msg
                Call WriteVar(f, "INFO", "CantMSG", 1)
            End If
            Close #n2
            End If
            Exit Sub
    End Select
    
    
    Select Case UCase$(Left$(rData, 6))
        Case "DESPHE" 'Mover Hechizo de lugar
            rData = Right(rData, Len(rData) - 6)
            Call DesplazarHechizo(UserIndex, CInt(ReadField(1, rData, 44)), CInt(ReadField(2, rData, 44)))
            Exit Sub
        Case "DESCOD" 'Informacion del hechizo
                rData = Right$(rData, Len(rData) - 6)
                Call modGuilds.ActualizarCodexYDesc(rData, UserList(UserIndex).GuildIndex)
                Exit Sub
    End Select
    
    '[Alejo]
    Select Case UCase$(Left$(rData, 7))
    Case "OFRECER"
            rData = Right$(rData, Len(rData) - 7)
            Arg1 = ReadField(1, rData, Asc(","))
            Arg2 = ReadField(2, rData, Asc(","))

            If val(Arg1) <= 0 Or val(Arg2) <= 0 Then
                Exit Sub
            End If
            If UserList(UserList(UserIndex).ComUsu.DestUsu).flags.UserLogged = False Then
                'sigue vivo el usuario ?
                Call FinComerciarUsu(UserIndex)
                Exit Sub
            Else
                'esta vivo ?
                If UserList(UserList(UserIndex).ComUsu.DestUsu).flags.Muerto = 1 Then
                    Call FinComerciarUsu(UserIndex)
                    Exit Sub
                End If
                '//Tiene la cantidad que ofrece ??//'
                If val(Arg1) = FLAGORO Then
                    'oro
                    If val(Arg2) > UserList(UserIndex).Stats.GLD Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No tienes esa cantidad." & FONTTYPE_TALK)
                        Exit Sub
                    End If
                Else
                    'inventario
                    If val(Arg2) > UserList(UserIndex).Invent.Object(val(Arg1)).Amount Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No tienes esa cantidad." & FONTTYPE_TALK)
                        Exit Sub
                    End If
                End If
                If UserList(UserIndex).ComUsu.Objeto > 0 Then
                    Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No puedes cambiar tu oferta." & FONTTYPE_TALK)
                    Exit Sub
                End If
                'No permitimos vender barcos mientras están equipados (no podés desequiparlos y causa errores)
                If UserList(UserIndex).flags.Navegando = 1 Then
                    If UserList(UserIndex).Invent.BarcoSlot = val(Arg1) Then
                        Call SendData(SendTarget.ToIndex, UserIndex, 0, "||No podés vender tu barco mientras lo estés usando." & FONTTYPE_TALK)
                        Exit Sub
                    End If
                End If
                
                UserList(UserIndex).ComUsu.Objeto = val(Arg1)
                UserList(UserIndex).ComUsu.cant = val(Arg2)
                If UserList(UserList(UserIndex).ComUsu.DestUsu).ComUsu.DestUsu <> UserIndex Then
                    Call FinComerciarUsu(UserIndex)
                    Exit Sub
                Else
                    '[CORREGIDO]
                    If UserList(UserList(UserIndex).ComUsu.DestUsu).ComUsu.Acepto = True Then
                        'NO NO NO vos te estas pasando de listo...
                        UserList(UserList(UserIndex).ComUsu.DestUsu).ComUsu.Acepto = False
                        Call SendData(SendTarget.ToIndex, UserList(UserIndex).ComUsu.DestUsu, 0, "||" & UserList(UserIndex).name & " ha cambiado su oferta." & FONTTYPE_TALK)
                    End If
                    '[/CORREGIDO]
                    'Es la ofrenda de respuesta :)
                    Call EnviarObjetoTransaccion(UserList(UserIndex).ComUsu.DestUsu)
                End If
            End If
            Exit Sub
    End Select
    '[/Alejo]
    
    Select Case UCase$(Left$(rData, 8))
        'clanesnuevo
        Case "ACEPPEAT" 'aceptar paz
            rData = Right$(rData, Len(rData) - 8)
            tInt = modGuilds.r_AceptarPropuestaDePaz(UserIndex, rData, tStr)
            If tInt = 0 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & tStr & FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.ToGuildMembers, UserList(UserIndex).GuildIndex, 0, "||Tu clan ha firmado la paz con " & rData & FONTTYPE_GUILD)
                Call SendData(SendTarget.ToGuildMembers, tInt, 0, "||Tu clan ha firmado la paz con " & UserList(UserIndex).name & FONTTYPE_GUILD)
            End If
            Exit Sub
        Case "RECPALIA" 'rechazar alianza
            rData = Right$(rData, Len(rData) - 8)
            tInt = modGuilds.r_RechazarPropuestaDeAlianza(UserIndex, rData, tStr)
            If tInt = 0 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & tStr & FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.ToGuildMembers, UserList(UserIndex).GuildIndex, 0, "||Tu clan rechazado la propuesta de alianza de " & rData & FONTTYPE_GUILD)
                Call SendData(SendTarget.ToGuildMembers, tInt, 0, "||" & UserList(UserIndex).name & " ha rechazado nuestra propuesta de alianza con su clan." & FONTTYPE_GUILD)
            End If
            Exit Sub
        Case "RECPPEAT" 'rechazar propuesta de paz
            rData = Right$(rData, Len(rData) - 8)
            tInt = modGuilds.r_RechazarPropuestaDePaz(UserIndex, rData, tStr)
            If tInt = 0 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & tStr & FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.ToGuildMembers, UserList(UserIndex).GuildIndex, 0, "||Tu clan rechazado la propuesta de paz de " & rData & FONTTYPE_GUILD)
                Call SendData(SendTarget.ToGuildMembers, tInt, 0, "||" & UserList(UserIndex).name & " ha rechazado nuestra propuesta de paz con su clan." & FONTTYPE_GUILD)
            End If
            Exit Sub
        Case "ACEPALIA" 'aceptar alianza
            rData = Right$(rData, Len(rData) - 8)
            tInt = modGuilds.r_AceptarPropuestaDeAlianza(UserIndex, rData, tStr)
            If tInt = 0 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & tStr & FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.ToGuildMembers, UserList(UserIndex).GuildIndex, 0, "||Tu clan ha firmado la alianza con " & rData & FONTTYPE_GUILD)
                Call SendData(SendTarget.ToGuildMembers, tInt, 0, "||Tu clan ha firmado la paz con " & UserList(UserIndex).name & FONTTYPE_GUILD)
            End If
            Exit Sub
        Case "PEACEOFF"
            'un clan solicita propuesta de paz a otro
            rData = Right$(rData, Len(rData) - 8)
            Arg1 = ReadField(1, rData, Asc(","))
            Arg2 = ReadField(2, rData, Asc(","))
            If modGuilds.r_ClanGeneraPropuesta(UserIndex, Arg1, PAZ, Arg2, Arg3) Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Propuesta de paz enviada" & FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & Arg3 & FONTTYPE_GUILD)
            End If
            Exit Sub
        Case "ALLIEOFF" 'un clan solicita propuesta de alianza a otro
            rData = Right$(rData, Len(rData) - 8)
            Arg1 = ReadField(1, rData, Asc(","))
            Arg2 = ReadField(2, rData, Asc(","))
            If modGuilds.r_ClanGeneraPropuesta(UserIndex, Arg1, ALIADOS, Arg2, Arg3) Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Propuesta de alianza enviada" & FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & Arg3 & FONTTYPE_GUILD)
            End If
            Exit Sub
        Case "ALLIEDET"
            'un clan pide los detalles de una propuesta de ALIANZA
            rData = Right$(rData, Len(rData) - 8)
            tStr = modGuilds.r_VerPropuesta(UserIndex, rData, ALIADOS, Arg1)
            If tStr = vbNullString Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & Arg1 & FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "ALLIEDE" & tStr)
            End If
            Exit Sub
        Case "PEACEDET" '-"ALLIEDET"
            'un clan pide los detalles de una propuesta de paz
            rData = Right$(rData, Len(rData) - 8)
            tStr = modGuilds.r_VerPropuesta(UserIndex, rData, PAZ, Arg1)
            If tStr = vbNullString Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & Arg1 & FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "PEACEDE" & tStr)
            End If
            Exit Sub
        Case "ENVCOMEN"
            rData = Trim$(Right$(rData, Len(rData) - 8))
            If rData = vbNullString Then Exit Sub
            tStr = modGuilds.a_DetallesAspirante(UserIndex, rData)
            If tStr = vbNullString Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "|| El personaje no ha mandado solicitud, o no estás habilitado para verla." & FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "PETICIO" & tStr)
            End If
            Exit Sub
        Case "ENVALPRO" 'enviame la lista de propuestas de alianza
            tIndex = modGuilds.r_CantidadDePropuestas(UserIndex, ALIADOS)
            tStr = "ALLIEPR" & tIndex & ","
            If tIndex > 0 Then
                tStr = tStr & modGuilds.r_ListaDePropuestas(UserIndex, ALIADOS)
            End If
            Call SendData(SendTarget.ToIndex, UserIndex, 0, tStr)
            Exit Sub
        Case "ENVPROPP" 'enviame la lista de propuestas de paz
            tIndex = modGuilds.r_CantidadDePropuestas(UserIndex, PAZ)
            tStr = "PEACEPR" & tIndex & ","
            If tIndex > 0 Then
                tStr = tStr & modGuilds.r_ListaDePropuestas(UserIndex, PAZ)
            End If
            Call SendData(SendTarget.ToIndex, UserIndex, 0, tStr)
            Exit Sub
        Case "DECGUERR" 'declaro la guerra
            rData = Right$(rData, Len(rData) - 8)
            tInt = modGuilds.r_DeclararGuerra(UserIndex, rData, tStr)
            If tInt = 0 Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & tStr & FONTTYPE_GUILD)
            Else
                'WAR shall be!
                Call SendData(SendTarget.ToGuildMembers, UserList(UserIndex).GuildIndex, 0, "|| TU CLAN HA ENTRADO EN GUERRA CON " & rData & FONTTYPE_GUILD)
                Call SendData(SendTarget.ToGuildMembers, tInt, 0, "|| " & UserList(UserIndex).name & " LE DECLARA LA GUERRA A TU CLAN" & FONTTYPE_GUILD)
            End If
            Exit Sub
        Case "NEWWEBSI"
            rData = Right$(rData, Len(rData) - 8)
            Call modGuilds.ActualizarWebSite(UserIndex, rData)
            Exit Sub
        Case "ACEPTARI"
            rData = Right$(rData, Len(rData) - 8)
            If Guilds(UserList(UserIndex).GuildIndex).CantidadDeMiembros >= 15 Then
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "||El clan esta lleno." & FONTTYPE_GUILD)
            Exit Sub
            End If
            If Not modGuilds.a_AceptarAspirante(UserIndex, rData, tStr) Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & tStr & FONTTYPE_GUILD)
            Else
                tInt = NameIndex(rData)
                If tInt > 0 Then
                    Call modGuilds.m_ConectarMiembroAClan(tInt, UserList(UserIndex).GuildIndex)
                End If
                Call SendData(SendTarget.ToGuildMembers, UserList(UserIndex).GuildIndex, 0, "||" & rData & " ha sido aceptado como miembro del clan." & FONTTYPE_GUILD)
            End If
            Exit Sub
        Case "RECHAZAR"
            rData = Trim$(Right$(rData, Len(rData) - 8))
            Arg1 = ReadField(1, rData, Asc(","))
            Arg2 = ReadField(2, rData, Asc(","))
            If Not modGuilds.a_RechazarAspirante(UserIndex, Arg1, Arg2, Arg3) Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "|| " & Arg3 & FONTTYPE_GUILD)
            Else
                tInt = NameIndex(Arg1)
                tStr = Arg3 & ": " & Arg2       'el mensaje de rechazo
                If tInt > 0 Then
                    Call SendData(SendTarget.ToIndex, tInt, 0, "|| " & tStr & FONTTYPE_GUILD)
                Else
                    'hay que grabar en el char su rechazo
                    Call modGuilds.a_RechazarAspiranteChar(Arg1, UserList(UserIndex).GuildIndex, Arg2)
                End If
            End If
            Exit Sub
        Case "ECHARCLA"
            'el lider echa de clan a alguien
            rData = Trim$(Right$(rData, Len(rData) - 8))
            tInt = modGuilds.m_EcharMiembroDeClan(UserIndex, rData)
            If tInt > 0 Then
                Call SendData(SendTarget.ToGuildMembers, tInt, 0, "||" & rData & " fue expulsado del clan." & FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "|| No puedes expulsar ese personaje del clan." & FONTTYPE_GUILD)
            End If
            Exit Sub
        Case "ACTGNEWS"
            rData = Right$(rData, Len(rData) - 8)
            Call modGuilds.ActualizarNoticias(UserIndex, rData)
            Exit Sub
        Case "1HRINFO<"
            rData = Right$(rData, Len(rData) - 8)
            If Trim$(rData) = vbNullString Then Exit Sub
            tStr = modGuilds.a_DetallesPersonaje(UserIndex, rData, Arg1)
            If tStr = vbNullString Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & Arg1 & FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "CHRINFO" & tStr)
            End If
            Exit Sub
        Case "ABREELEC"
            If Not modGuilds.v_AbrirElecciones(UserIndex, tStr) Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & tStr & FONTTYPE_GUILD)
            Else
                Call SendData(SendTarget.ToGuildMembers, UserList(UserIndex).GuildIndex, 0, "||¡Han comenzado las elecciones del clan! Puedes votar escribiendo /VOTO seguido del nombre del personaje, por ejemplo: /VOTO " & UserList(UserIndex).name & FONTTYPE_GUILD)
            End If
            Exit Sub
    End Select
    

    Select Case UCase$(Left$(rData, 9))
        Case "SOLICITUD"
             rData = Right$(rData, Len(rData) - 9)
             Arg1 = ReadField(1, rData, Asc(","))
             Arg2 = ReadField(2, rData, Asc(","))
             If Not modGuilds.a_NuevoAspirante(UserIndex, Arg1, Arg2, tStr) Then
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||" & tStr & FONTTYPE_GUILD)
             Else
                Call SendData(SendTarget.ToIndex, UserIndex, 0, "||Tu solicitud ha sido enviada. Espera prontas noticias del líder de " & Arg1 & "." & FONTTYPE_GUILD)
             End If
             Exit Sub
    End Select
    
    Select Case UCase$(Left$(rData, 11))
        Case "CLANDETAILS"
            rData = Right$(rData, Len(rData) - 11)
            If Trim$(rData) = vbNullString Then Exit Sub
            Call SendData(SendTarget.ToIndex, UserIndex, 0, "CLANDET" & modGuilds.SendGuildDetails(rData))
            Exit Sub
    End Select
    
Procesado = False
    
End Sub
