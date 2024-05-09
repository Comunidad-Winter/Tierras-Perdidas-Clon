Attribute VB_Name = "Balance"
'///////// DECLARACIONES /////////////////////
Public MagoElfo                As Integer
Public MagoHumano              As Integer
Public MagoGnomo               As Integer
Public MagoElfoOscuro          As Integer
Public MagoEnano               As Integer

Public AsesinoElfo             As Integer
Public AsesinoHumano           As Integer
Public AsesinoElfoOscuro       As Integer
Public AsesinoEnano            As Integer
Public AsesinoGnomo            As Integer

Public ClerigoHumano           As Integer
Public ClerigoElfo             As Integer
Public ClerigoElfoOscuro       As Integer
Public ClerigoEnano            As Integer
Public ClerigoGnomo            As Integer

Public BardoHumano             As Integer
Public BardoElfo               As Integer
Public BardoElfoOscuro         As Integer
Public BardoEnano              As Integer
Public BardoGnomo              As Integer

Public GuerreroHumano          As Integer
Public GuerreroElfo            As Integer
Public GuerreroElfoOscuro      As Integer
Public GuerreroGnomo           As Integer
Public GuerreroEnano           As Integer

Public CazadorHumano           As Integer
Public CazadorElfo             As Integer
Public CazadorGnomo            As Integer
Public CazadorEnano            As Integer
Public CazadorElfoOscuro       As Integer

Public DruidaHumano            As Integer
Public DruidaElfo              As Integer
Public DruidaElfoOscuro        As Integer
Public DruidaGnomo             As Integer
Public DruidaEnano             As Integer

Public PaladinEnano            As Integer
Public PaladinHumano           As Integer
Public PaladinElfoOscuro       As Integer
Public PaladinElfo             As Integer
Public PaladinGnomo            As Integer

Public Type Evaciones

    Asesino                    As String
    Bardo                      As String
    Ladron                     As String
    Paladin                    As String
    Cazador                    As String
    Guerrero                   As String
    Others                     As String
    
End Type

Public Type AtaqueArmas

    Asesino                    As String
    Bardo                      As String
    Ladron                     As String
    Paladin                    As String
    Cazador                    As String
    Guerrero                   As String
    Others                     As String
    Pirata                     As String
    Bandido                    As String
    Clerigo                    As String
    Druida                     As String
    
End Type

Public otAtaqueArmas           As AtaqueArmas
Public otEvaciones             As Evaciones
Public BalancePath             As String
Public AtcArmaPath             As String
Public EvacionPath             As String

'/////////////////////////////////////////////

'///////// SUBS Y FUNCTIONS ///////////////////
Public Function InitBalance()

 BalancePath = App.Path & "\Balances\Balance_Vidas.ini"

 MagoElfo = GetVar(BalancePath, "CLASES", "MagoElfo")
 MagoHumano = GetVar(BalancePath, "CLASES", "MagoHumano")
 MagoGnomo = GetVar(BalancePath, "CLASES", "MagoGnomo")
 MagoElfoOscuro = GetVar(BalancePath, "CLASES", "MagoElfoOscuro")
 MagoEnano = GetVar(BalancePath, "CLASES", "MagoEnano")

 AsesinoElfo = GetVar(BalancePath, "CLASES", "AsesinoElfo")
 AsesinoHumano = GetVar(BalancePath, "CLASES", "AsesinoHumano")
 AsesinoElfoOscuro = GetVar(BalancePath, "CLASES", "AsesinoElfoOscuro")
 AsesinoEnano = GetVar(BalancePath, "CLASES", "AsesinoEnano")
 AsesinoGnomo = GetVar(BalancePath, "CLASES", "AsesinoGnomo")

 ClerigoHumano = GetVar(BalancePath, "CLASES", "ClerigoHumano")
 ClerigoElfo = GetVar(BalancePath, "CLASES", "ClerigoElfo")
 ClerigoElfoOscuro = GetVar(BalancePath, "CLASES", "ClerigoElfoOscuro")
 ClerigoEnano = GetVar(BalancePath, "CLASES", "ClerigoEnano")
 ClerigoGnomo = GetVar(BalancePath, "CLASES", "ClerigoGnomo")

 BardoHumano = GetVar(BalancePath, "CLASES", "BardoHumano")
 BardoElfo = GetVar(BalancePath, "CLASES", "BardoElfo")
 BardoElfoOscuro = GetVar(BalancePath, "CLASES", "BardoElfoOscuro")
 BardoEnano = GetVar(BalancePath, "CLASES", "BardoEnano")
 BardoGnomo = GetVar(BalancePath, "CLASES", "BardoGnomo")

 GuerreroHumano = GetVar(BalancePath, "CLASES", "GuerreroHumano")
 GuerreroElfo = GetVar(BalancePath, "CLASES", "GuerreroElfo")
 GuerreroElfoOscuro = GetVar(BalancePath, "CLASES", "GuerreroElfoOscuro")
 GuerreroGnomo = GetVar(BalancePath, "CLASES", "GuerreroGnomo")
 GuerreroEnano = GetVar(BalancePath, "CLASES", "GuerreroEnano")

 CazadorHumano = GetVar(BalancePath, "CLASES", "CazadorHumano")
 CazadorElfo = GetVar(BalancePath, "CLASES", "CazadorElfo")
 CazadorElfoOscuro = GetVar(BalancePath, "CLASES", "CazadorElfoOscuro")
 CazadorGnomo = GetVar(BalancePath, "CLASES", "CazadorGnomo")
 CazadorEnano = GetVar(BalancePath, "CLASES", "CazadorEnano")

 DruidaHumano = GetVar(BalancePath, "CLASES", "DruidaHumano")
 DruidaElfo = GetVar(BalancePath, "CLASES", "DruidaElfo")
 DruidaElfoOscuro = GetVar(BalancePath, "CLASES", "DruidaElfoOscuro")
 DruidaGnomo = GetVar(BalancePath, "CLASES", "DruidaGnomo")
 DruidaEnano = GetVar(BalancePath, "CLASES", "DruidaEnano")

 PaladinEnano = GetVar(BalancePath, "CLASES", "PaladinEnano")
 PaladinHumano = GetVar(BalancePath, "CLASES", "PaladinHumano")
 PaladinElfoOscuro = GetVar(BalancePath, "CLASES", "PaladinElfoOscuro")
 PaladinElfo = GetVar(BalancePath, "CLASES", "PaladinElfo")
 PaladinGnomo = GetVar(BalancePath, "CLASES", "PaladinGnomo")
 
End Function

Public Function InitEvaciones()

 EvacionPath = App.Path & "\Balances\Balance_Evasiones.ini"

 With otEvaciones

 .Asesino = GetVar(EvacionPath, "Asesino", "Evacion")
 .Bardo = GetVar(EvacionPath, "Bardo", "Evacion")
 .Cazador = GetVar(EvacionPath, "Cazador", "Evacion")
 .Paladin = GetVar(EvacionPath, "Paladin", "Evacion")
 .Ladron = GetVar(EvacionPath, "Ladron", "Evacion")
 .Guerrero = GetVar(EvacionPath, "Guerrero", "Evacion")
 .Others = GetVar(EvacionPath, "Otros", "Evacion")
 
 End With
 
End Function

Public Function InitGolpesArmas()

 AtcArmaPath = App.Path & "\Balances\Balance_Ataques.ini"

 With otAtaqueArmas
 
 .Asesino = GetVar(AtcArmaPath, "ASESINO", "ModificadorPoderAtaqueArmas")
 .Bandido = GetVar(AtcArmaPath, "BANDIDO", "ModificadorPoderAtaqueArmas")
 .Bardo = GetVar(AtcArmaPath, "BARDO", "ModificadorPoderAtaqueArmas")
 .Cazador = GetVar(AtcArmaPath, "CAZADOR", "ModificadorPoderAtaqueArmas")
 .Clerigo = GetVar(AtcArmaPath, "CLERIGO", "ModificadorPoderAtaqueArmas")
 .Druida = GetVar(AtcArmaPath, "DRUIDA", "ModificadorPoderAtaqueArmas")
 .Guerrero = GetVar(AtcArmaPath, "GUERRERO", "ModificadorPoderAtaqueArmas")
 .Ladron = GetVar(AtcArmaPath, "LADRON", "ModificadorPoderAtaqueArmas")
 .Paladin = GetVar(AtcArmaPath, "PALADIN", "ModificadorPoderAtaqueArmas")
 .Pirata = GetVar(AtcArmaPath, "PIRATA", "ModificadorPoderAtaqueArmas")
 .Others = GetVar(AtcArmaPath, "OTROS", "ModificadorPoderAtaqueArmas")
 
 End With
 
End Function

