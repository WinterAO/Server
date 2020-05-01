Attribute VB_Name = "ModBOTS"
Option Explicit
 
'Defensa del bot jeje
Private Const IA_MINDEF  As Integer = 10
Private Const IA_MAXDEF  As Integer = 12
 
Public MAX_BOTS    As Byte
 
'Charindex reservado.
Private IA_CHAR     As Integer
 
'Cantidad de hechizos que lanza
 
Private Const IA_M_SPELL As Byte = 3
Private Const IA_NUMCHAT As Byte = 5
 
'Constantes de intervalos.
 
Private Const IA_SINT   As Integer = 800    'Intervalo entre hechizo-hechizo.
Private Const IA_SREMO  As Integer = 500    'Intervalo remo.
Private Const IA_MOVINT As Integer = 240    'Intervalo caminta.
Private Const IA_USEOBJ As Integer = 200    'Intervalo usar potas.
Private Const IA_HITINT As Integer = 200    'Intervalo para golpe
Private Const IA_PROINT As Integer = 700    'Intervalo de flecha
Private Const IA_TALKIN As Integer = 4000   'Intervalo de hablAR :P
 
'Probabilidades de que te pegue
 
Private Const IA_CASTEO As Byte = 77
 
Private Const IA_PROBEV As Byte = 160
Private Const IA_PROBEX As Byte = 220
 
Private Const IA_SLOTS  As Byte = 20
 
Type ia_Interval
     SpellCount         As Byte         'Intervalo para tirar hechizos.
     UseItemCount       As Byte         'Intervalo para usar pociones.
     MoveCharCount      As Byte         'Intervalo para mover el char.
     ParalizisCount     As Byte         'Intervalo para removerse.
     HitCount           As Byte         'Intervalo para pegar golpesito.
     ArrowCount         As Byte         'Intervalo para flechas
     ChatCount          As Byte         'INtervalo para hablar XD
End Type
 
Type ia_Spells
     DamageMin          As Byte         'Minimo daño que hace.
     DamageMax          As Byte         'Maximo daño que hace.
     spellIndex         As Byte         'Lo usamos para el fx.
End Type
 
Enum eIASupportActions
     SRemover = 1                       'Remueve.
     SCurar = 2                         'Cura.
End Enum
 
Enum eIAClase
     Clerigo = 1                        'Bot Clero
     Mago = 2                           'Bot Mago
     Cazador = 3                        'Bot kza
End Enum
 
Enum eIAactions
     ePegar = 1                          'accion pegar.
     eMagia = 2                          'atacar con hechizo
End Enum
 
Enum eIAMoviments
     SeguirVictima = 1                   'Si seguia la victima
     MoverRandom = 2                     'Random moviment :P
End Enum
 
Type BOT
     Name               As String       'Nombre del BOT
     Level              As Byte         'Nivel del BOT
     Raza               As eRaza        'Raza del BOT
     Genero             As eGenero      'Genero del BOT
     EsCriminal         As Boolean
     Pos                As WorldPos     'Posicion en el mundo.
     maxVida            As Integer      'Maxima vida.
     minVida            As Integer      'Vida del bot.
     clase              As eIAClase     'Clases de bot.
     minMana            As Integer   'Mana del bot.
     maxMana            As Integer      'Maxima mana
     Char               As Char         'Apariencia.
     Invocado           As Boolean      'Si existe.
     Paralizado         As Boolean      'Si está inmo.
     Intervalos         As ia_Interval  'Intervalos de acciones.
     Viajante           As Boolean      'Bot Viajante :P
     ViajanteUser       As Integer      'Usuario que atacó al viajante.
     UltimaAccion       As eIAactions   'ULTIMA ACCION/ATAQUE.
     UltimoMovimiento   As eIAMoviments 'ULTIMO MOVIMIENTO
     Navegando          As Boolean      'Navegando?
     ViajanteAntes      As WorldPos     'Pos cuando un viajante ataca un usuario.
     NroItems           As Byte         'Numero de Items en el inventario
     Inv(1 To IA_SLOTS) As obj          'Inventario del bot.
     UltimaIdaObjeto    As Boolean      'Ultimo movimiento fue buscar objs?
End Type
 
Public IA_Bot()                       As BOT
Public IA_spell(1 To IA_M_SPELL)      As ia_Spells
Public IA_Chats(1 To IA_NUMCHAT)      As String
 
'Cantidad de bots invocados.
Public NumInvocados                    As Byte
 
Function IA_EquiparCasco(ByVal BotIndex As Byte) As Integer
'****************************************************
'Autor: Lorwik
'Fecha 01/05/2020
'Descripcion: Si el BOT tiene un casco en el inventario lo equipamos
'****************************************************
    Dim i As Byte
    
    With IA_Bot(BotIndex)

        '¿Tiene items en el inventario?
        If .NroItems > 0 Then
            For i = 1 To .NroItems
                'Vamos a equipar el primer casco que encuentre
                If ObjData(.Inv(i).ObjIndex).OBJType = otCasco Then
                    IA_EquiparCasco = .Inv(i).ObjIndex
                Else
                    IA_EquiparCasco = -1
                End If
            
            Next i
        Else
            IA_EquiparCasco = -1
        End If
        
    End With
 
End Function
 
Function IA_EquiparEscudo(ByVal BotIndex As Byte) As Integer
'****************************************************
'Autor: Lorwik
'Fecha 01/05/2020
'Descripcion: Si el BOT tiene un escudo en el inventario lo equipamos
'****************************************************
    Dim i As Byte
    
    With IA_Bot(BotIndex)

        '¿Tiene items en el inventario?
        If .NroItems > 0 Then
            For i = 1 To .NroItems
                'Vamos a equipar el primer escudo que encuentre
                If ObjData(.Inv(i).ObjIndex).OBJType = otEscudo Then
                    IA_EquiparEscudo = .Inv(i).ObjIndex
                Else
                    IA_EquiparEscudo = -1
                End If
            
            Next i
        Else
            IA_EquiparEscudo = -1
        End If
        
    End With
 
End Function
 
Function IA_EquiparArma(ByVal BotIndex As Byte) As Integer
'****************************************************
'Autor: Lorwik
'Fecha 01/05/2020
'Descripcion: Si el BOT tiene un escudo en el inventario lo equipamos
'****************************************************
    Dim i As Byte
    
    With IA_Bot(BotIndex)

        '¿Tiene items en el inventario?
        If .NroItems > 0 Then
            For i = 1 To .NroItems
                'Vamos a equipar el primer arma que encuentre
                If ObjData(.Inv(i).ObjIndex).OBJType = otWeapon Then
                    IA_EquiparArma = .Inv(i).ObjIndex
                Else
                    IA_EquiparArma = -1
                End If
            
            Next i
        Else
            IA_EquiparArma = -1
        End If
        
    End With
 
End Function

Function IA_EquiparArmadura(ByVal BotIndex As Byte) As Integer
'****************************************************
'Autor: Lorwik
'Fecha 01/05/2020
'Descripcion: Si el BOT tiene ropa en el inventario lo equipamos
'****************************************************
    Dim i As Byte
    
    With IA_Bot(BotIndex)

        '¿Tiene items en el inventario?
        If .NroItems > 0 Then
            For i = 1 To .NroItems
                'Vamos a equipar la primera armadura que encuentre
                If ObjData(.Inv(i).ObjIndex).OBJType = otArmadura Then
                    IA_EquiparArmadura = .Inv(i).ObjIndex
                Else
                    IA_EquiparArmadura = -1
                End If
            
            Next i
        Else
            IA_EquiparArmadura = -1
        End If
        
    End With
 
End Function
 
Function ia_CalcularGolpe(ByVal victimIndex As Integer) As Integer
 
    ' @designer     :  maTih.-
    ' @date         :  2012/02/01
    ' @note         :  Calcula el golpe (daño) q hace el bot al user.
     
    Dim ParteCuerpo     As Integer
    Dim DañoAbsorvido   As Integer
     
    ParteCuerpo = RandomNumber(PartesCuerpo.bCabeza, PartesCuerpo.bTorso)
     
    'Si pega en la cabeza.
    If ParteCuerpo = PartesCuerpo.bCabeza Then
       'Si tiene casco baja el golpe
           If UserList(victimIndex).Invent.CascoEqpObjIndex <> 0 Then
              DañoAbsorvido = RandomNumber(ObjData(UserList(victimIndex).Invent.CascoEqpObjIndex).MinDef, ObjData(UserList(victimIndex).Invent.CascoEqpObjIndex).MaxDef)
           End If
    Else
        'Se fija por la armadura.
           If UserList(victimIndex).Invent.ArmourEqpObjIndex <> 0 Then
              DañoAbsorvido = RandomNumber(ObjData(UserList(victimIndex).Invent.ArmourEqpObjIndex).MinDef, ObjData(UserList(victimIndex).Invent.ArmourEqpObjIndex).MaxDef)
           End If
    End If
           
    'DEVUELVE.
    ia_CalcularGolpe = (RandomNumber(150, 180) - DañoAbsorvido)
       
End Function
 
Function ia_AciertaGolpe(ByVal victimIndex As Integer) As Boolean
 
    ' @designer     :  maTih.-
    ' @date         :  2012/02/01
    ' @note         :  Evasión del usuario, esto le faltan unos retoques.
     
    Dim tempEvasion     As Long
    Dim tempEvasionEsc  As Long
    Dim tempResultado   As Long
     
    'Evasión del usuario.
    tempEvasion = PoderEvasion(victimIndex)
     
    'Evasión del usuario con escudos.
    'Tiene escudo?
    If UserList(victimIndex).Invent.EscudoEqpObjIndex <> 0 Then
        tempEvasionEsc = PoderEvasionEscudo(victimIndex)
        tempEvasionEsc = tempEvasion + tempEvasionEsc
    Else
        tempEvasionEsc = 0
    End If
     
    'Acierta?
    tempResultado = MaximoInt(10, MinimoInt(90, 50 + (IA_PROBEX - tempEvasion) * 0.4))
     
    'Random.
    ia_AciertaGolpe = (RandomNumber(1, 100) <= tempResultado)
 
End Function
 
Function ia_PuedeMeele(ByRef PosBot As WorldPos, ByRef PosVictim As WorldPos, ByRef NewHeading As eHeading) As Boolean
 
    ' @designer     :  maTih.-
    ' @date         :  2012/02/01
    ' @note         :  Se fija si está al lado, y guarda el heading.
     
    With PosVictim
        
        'Mirando hacia la derecha lo tiene ?
        If PosBot.X + 1 = .X Then
           ia_PuedeMeele = (.Y = PosBot.Y)
           
           If ia_PuedeMeele Then
              NewHeading = eHeading.EAST
           End If
           
           Exit Function
        End If
        
        'mirando hacia izq?
        If PosBot.X - 1 = .X Then
           ia_PuedeMeele = (.Y = PosBot.Y)
           
           If ia_PuedeMeele Then
              NewHeading = eHeading.WEST
           End If
           
           Exit Function
        End If
        
        'mirando arriba
        If PosBot.Y - 1 = .Y Then
           ia_PuedeMeele = (.X = PosBot.X)
           
           If ia_PuedeMeele Then
              NewHeading = eHeading.NORTH
           End If
           
           Exit Function
        End If
        
        'Abajo.
        If PosBot.Y + 1 = .Y Then
           ia_PuedeMeele = (PosBot.X = .X)
           
           If ia_PuedeMeele Then
              NewHeading = eHeading.SOUTH
           End If
           
           Exit Function
        End If
        
    End With
 
End Function
 
Sub ia_CreateChar(ByVal ProximoBot As Byte)
'******************************************
'Autor: maTih.-
'Fecha:: 2012/02/01
'Descripcion: Crea el char.
'Lorwik> Ahora comprueba si tiene arma, escudo y casco, y lo equipa
'******************************************

    Dim PackageToSend   As String
    Dim Armadura        As Integer
    Dim Arma            As Integer
    Dim Escudo          As Integer
    Dim Casco           As Integer

    With IA_Bot(ProximoBot).Char
        
        Armadura = IA_EquiparArmadura(ProximoBot)
        If Armadura <> -1 Then
            .Body = ObjData(Armadura).Ropaje
        Else 'Si no tiene le damos un cuerpo desnudo
            .Body = CuerpoDesnudo(IA_Bot(ProximoBot).Genero, IA_Bot(ProximoBot).Raza)
        End If
    
        '¿Tiene Arma?
        Arma = IA_EquiparArma(ProximoBot)
        If Arma <> -1 Then .WeaponAnim = ObjData(Arma).WeaponAnim
        
        '¿Tiene escudo?
        Escudo = IA_EquiparEscudo(ProximoBot)
        If Escudo <> -1 Then .ShieldAnim = ObjData(Escudo).ShieldAnim
            
        '¿Tiene Casco?
        Casco = IA_EquiparCasco(ProximoBot)
        If Casco <> -1 Then .CascoAnim = ObjData(Casco).CascoAnim
        
        'Precalculado : P
        .CharIndex = IA_CHAR + ProximoBot
        
        'Preparo el paquete de datos.
        
        Dim tmp_Color   As eNickColor
                
        If IA_Bot(ProximoBot).EsCriminal Then
            tmp_Color = eNickColor.ieCriminal
        Else
            tmp_Color = eNickColor.ieCiudadano
        End If
        
        PackageToSend = PrepareMessageCharacterCreate(.Body, .Head, eHeading.SOUTH, .CharIndex, IA_Bot(ProximoBot).Pos.X, IA_Bot(ProximoBot).Pos.Y, .WeaponAnim, .ShieldAnim, 0, 0, .CascoAnim, IA_Bot(ProximoBot).Name, tmp_Color, 0, 0)
        
        'Actualizo el area.
        ia_SendToBotArea ProximoBot, PackageToSend
        
    End With
 
End Sub
 
Public Function ia_Spawn(ByRef PosToSpawn As WorldPos) As Integer
 
    ' @designer     :  maTih.-
    ' @date         :  2012/02/01
     
    Dim ProximoBot  As Byte
    Dim PackageSend As String
     
    ProximoBot = IA_GetNextSlot
     
    If Not ProximoBot <> 0 Then Exit Function
     
    With IA_Bot(ProximoBot)
        
        .Invocado = True
        
        .Paralizado = False
        
        'Seteo la posición.
        .Pos = PosToSpawn
        
        'Creo el char.
        ia_CreateChar ProximoBot
       
        'Primer action ! : D
        ia_Action ProximoBot
       
        'PackageSend = PrepareMessageChatOverHead("VeNGan PutOs xD!", .Char.CharIndex, vbCyan)
       
        'ia_SendToBotArea ProximoBot, PackageSend
       
        .Intervalos.SpellCount = 100
       
        NumInvocados = NumInvocados + 1
        
        MapData(.Pos.Map, .Pos.X, .Pos.Y).BotIndex = ProximoBot
        
        'devuelvo el id del bot
        ia_Spawn = ProximoBot
       
    End With
 
End Function
 
Public Sub ia_Spells()
 
    ' @designer     :  maTih.-
    ' @date         :  2012/02/01
     
    'Un poco hardcodeado pero bueno :D
     
    'Hechizo 1 : descarga.
    IA_spell(1).DamageMax = 120
    IA_spell(1).DamageMax = 177
    IA_spell(1).spellIndex = 23
     
    'Hechizo 2 : apoca
     
    IA_spell(2).DamageMin = 190
    IA_spell(2).DamageMax = 220
    IA_spell(2).spellIndex = 25
     
    'Paralizar.
    IA_spell(3).DamageMax = 0
    IA_spell(3).DamageMin = 0
    IA_spell(3).spellIndex = 9
     
End Sub
 
Sub ia_RandomMoveChar(ByVal BotIndex As Byte, ByVal siguiendoIndex As Integer, ByRef HError As Boolean)
 
    ' @designer     :  maTih.-
    ' @date         :  2012/02/01
     
    With IA_Bot(BotIndex)
     
        Dim nRandom     As Byte
       
        '25% De probabilidades de moverse a
        'cualquiera de las cuatro direcciones.
       
        nRandom = RandomNumber(1, 4)
       
        Select Case nRandom
       
               Case 1
               
                    If ia_LegalPos(.Pos.X + 1, .Pos.Y, BotIndex, siguiendoIndex) = False Then HError = True: Exit Sub
                    
                    'Borro el BotIndex del tile anterior.
                    MapData(.Pos.Map, .Pos.X, .Pos.Y).BotIndex = 0
                    .Pos.X = .Pos.X + 1
               
               Case 2
               
                    If ia_LegalPos(.Pos.X - 1, .Pos.Y, BotIndex, siguiendoIndex) = False Then HError = True: Exit Sub
               
                    'Borro el BotIndex del tile anterior.
                    MapData(.Pos.Map, .Pos.X, .Pos.Y).BotIndex = 0
                    .Pos.X = .Pos.X - 1
               
               Case 3
               
                    If ia_LegalPos(.Pos.X, .Pos.Y + 1, BotIndex, siguiendoIndex) = False Then HError = True: Exit Sub
               
                    'Borro el BotIndex del tile anterior.
                    MapData(.Pos.Map, .Pos.X, .Pos.Y).BotIndex = 0
                    .Pos.Y = .Pos.Y + 1
               
               Case 4
               
                    If ia_LegalPos(.Pos.X, .Pos.Y - 1, BotIndex, siguiendoIndex) = False Then HError = True: Exit Sub
                    
                    'Borro el BotIndex del tile anterior.
                    MapData(.Pos.Map, .Pos.X, .Pos.Y).BotIndex = 0
                    .Pos.Y = .Pos.Y - 1
       
        End Select
     
    End With
 
End Sub
 
Sub ia_CargarRutas(ByRef MAPFILE As String, ByVal MapIndex As Integer)
 
    ' @designer     :  maTih.-
    ' @date         :  2012/02/01
    ' @modificated  :  Carga las rutas de un mapa : D
     
    Dim loopX   As Long
    Dim loopY   As Long
    Dim tmpVal  As eHeading
     
    For loopX = 1 To 100
        For loopY = 1 To 100
            
            tmpVal = val(GetVar(MAPFILE, CStr(loopX) & "," & CStr(loopY), "Direccion"))
            
            If tmpVal <> 0 Then
              ' MapData(MapIndex, loopX, loopY).Rutas(1) = tmpVal
            End If
            
        Next loopY
    Next loopX
 
End Sub
 
Function ia_LegalPos(ByVal X As Byte, ByVal Y As Byte, ByVal BotIndex As Byte, Optional ByVal siguiendoUser As Integer = 0) As Boolean
 
    ' @designer     :  maTih.-
    ' @date         :  2012/02/01
    ' @modificated  :  Esta función ya no trabaja con la pos del npc si no que ahora usa los parámetros.
     
    ia_LegalPos = False
     
    With MapData(IA_Bot(BotIndex).Pos.Map, X, Y)
     
         '¿Es un mapa valido?
        If (IA_Bot(BotIndex).Pos.Map <= 0 Or IA_Bot(BotIndex).Pos.Map > NumMaps) Or (X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder) Then Exit Function
     
         'Tile bloqueado?
         If .Blocked <> 0 Then Exit Function
       
         'Hay un usuario?
         If .UserIndex > 0 Then
            'Si no es un adminInvisible entonces nos vamos.
            If UserList(.UserIndex).flags.AdminInvisible <> 1 Then Exit Function
        End If
     
        'Hay un NPC?
        If .NpcIndex <> 0 Then Exit Function
         
        'Hay un bot?
        If .BotIndex <> 0 Then Exit Function
        
        'Siguiendo Index?
        If siguiendoUser <> 0 Then
            'Válido para evitar el rango Y pero no su eje X.
            If Abs(Y - UserList(siguiendoUser).Pos.Y) > RANGO_VISION_Y Then Exit Function
       
            If Abs(X - UserList(siguiendoUser).Pos.X) > RANGO_VISION_X Then Exit Function
        End If
        
         ia_LegalPos = True
       
    End With
 
End Function
 
Sub ia_SearchPath(ByVal BotIndex As Byte, ByRef tPos As WorldPos, ByRef findHeading As eHeading)
 
    ' @designer     :  maTih.-
    ' @date         :  2012/03/13
    ' @                Buscá una ruta y devuelve un puntero con el heading.
     
    findHeading = FindDirection(IA_Bot(BotIndex).Pos, tPos)
 
End Sub
 
Sub ia_MoveToHeading(ByVal BotIndex As Byte, ByVal toHeading As eHeading, ByRef FoundErr As Boolean)
 
    ' @designer     :  maTih.-
    ' @date         :  2012/02/01
    ' @note         :  Mueve el char del npc hacia una posición.
     
    FoundErr = True
     
    Select Case toHeading
     
           Case eHeading.NORTH  '<Move norte.
                'No legal pos.
                If Not ia_LegalPos(IA_Bot(BotIndex).Pos.X, IA_Bot(BotIndex).Pos.Y - 1, BotIndex) Then Exit Sub
                
                'Se mueve, borro el anterior botIndex.
                MapData(IA_Bot(BotIndex).Pos.Map, IA_Bot(BotIndex).Pos.X, IA_Bot(BotIndex).Pos.Y).BotIndex = 0
                'Set la nueva posición
                IA_Bot(BotIndex).Pos.Y = IA_Bot(BotIndex).Pos.Y - 1
                
           Case eHeading.EAST   '<Move este.
                'Si hay posición inválida no se peude mover.
                If Not ia_LegalPos(IA_Bot(BotIndex).Pos.X + 1, IA_Bot(BotIndex).Pos.Y, BotIndex) Then Exit Sub
                
                'Se mueve, borro el anterior botIndex.
                MapData(IA_Bot(BotIndex).Pos.Map, IA_Bot(BotIndex).Pos.X, IA_Bot(BotIndex).Pos.Y).BotIndex = 0
                
                'Set la nueva posición
                IA_Bot(BotIndex).Pos.X = IA_Bot(BotIndex).Pos.X + 1
                
           Case eHeading.SOUTH  '<Move sur.
                'Si hay posición inválida no se peude mover.
                If Not ia_LegalPos(IA_Bot(BotIndex).Pos.X, IA_Bot(BotIndex).Pos.Y + 1, BotIndex) Then Exit Sub
                
                'Se mueve, borro el anterior botIndex.
                MapData(IA_Bot(BotIndex).Pos.Map, IA_Bot(BotIndex).Pos.X, IA_Bot(BotIndex).Pos.Y).BotIndex = 0
                
                'Set la nueva posición
                IA_Bot(BotIndex).Pos.Y = IA_Bot(BotIndex).Pos.Y + 1
                
           Case eHeading.WEST   '<Move oeste.
                'Si hay posición inválida no se peude mover.
                If Not ia_LegalPos(IA_Bot(BotIndex).Pos.X - 1, IA_Bot(BotIndex).Pos.Y, BotIndex) Then Exit Sub
                
                'Se mueve, borro el anterior botIndex.
                MapData(IA_Bot(BotIndex).Pos.Map, IA_Bot(BotIndex).Pos.X, IA_Bot(BotIndex).Pos.Y).BotIndex = 0
                
                'Set la nueva posición
                IA_Bot(BotIndex).Pos.X = IA_Bot(BotIndex).Pos.X - 1
                
    End Select
     
    FoundErr = False
 
End Sub
 
 
Sub ia_MoveViajante(ByVal BotIndex As Byte, ByVal Direccion As eHeading)
 
    ' @designer     :  maTih.-
    ' @date         :  2012/02/01
    ' @note         :  Move el viajante hacia una posición
     
    Dim HabiaAgua As Boolean
     
    With IA_Bot(BotIndex)
     
         'Hacia donde se mueve..
         Select Case Direccion
                
                Case eHeading.NORTH     'Norte.
                     MapData(.Pos.Map, .Pos.X, .Pos.Y).BotIndex = 0
                     .Pos.Y = .Pos.Y - 1
                     MapData(.Pos.Map, .Pos.X, .Pos.Y).BotIndex = BotIndex
                     
                Case eHeading.EAST      'Este.
                     MapData(.Pos.Map, .Pos.X, .Pos.Y).BotIndex = 0
                     .Pos.X = .Pos.X + 1
                     MapData(.Pos.Map, .Pos.X, .Pos.Y).BotIndex = BotIndex
                
                Case eHeading.SOUTH     'Sur.
                     MapData(.Pos.Map, .Pos.X, .Pos.Y).BotIndex = 0
                     .Pos.Y = .Pos.Y + 1
                     MapData(.Pos.Map, .Pos.X, .Pos.Y).BotIndex = BotIndex
                     
                Case eHeading.WEST      'Oeste.
                     MapData(.Pos.Map, .Pos.X, .Pos.Y).BotIndex = 0
                     .Pos.X = .Pos.X - 1
                     MapData(.Pos.Map, .Pos.X, .Pos.Y).BotIndex = BotIndex
         End Select
         
         HabiaAgua = HayAgua(.Pos.Map, .Pos.X, .Pos.Y)
         
         If HabiaAgua Then
            'Si hay agua cambio el cuerpo.
            ia_SendToBotArea BotIndex, PrepareMessageCharacterChange(395, 0, Direccion, .Char.CharIndex, 0, 0, 0, 0, 0)
            .Navegando = True
         Else
            'No habia agua, y... estaba navegando?
            If .Navegando Then
               'cambio el body y demas.
               ia_SendToBotArea BotIndex, PrepareMessageCharacterChange(.Char.Body, .Char.Head, Direccion, .Char.CharIndex, .Char.WeaponAnim, .Char.ShieldAnim, 0, 0, .Char.CascoAnim)
               .Navegando = False
            End If
        End If
        
         'Actualizamso
         ia_SendToBotArea BotIndex, PrepareMessageCharacterMove(.Char.CharIndex, .Pos.X, .Pos.Y)
            
    End With
 
End Sub
 
Function ia_HeadingToMolestNpc(ByVal NpcIndex As Integer) As eHeading
 
    ' @designer     :  maTih.-
    ' @date         :  2012/02/01
    ' @note         :  Devuelve un heading para un npc que está molestando el paso.
     
    Dim nPos    As WorldPos
     
    nPos = Npclist(NpcIndex).Pos
     
    With MapData(Npclist(NpcIndex).Pos.Map, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y)
     
         'Pos legal hacia arriba?
         If LegalPosNPC(nPos.Map, nPos.X, nPos.Y - 1, 0) Then
            'Mientras no halla bot.
            If Not .BotIndex <> 0 Then
               ia_HeadingToMolestNpc = eHeading.NORTH
            End If
         End If
         
         'Pos legal hacia abajo?
         If LegalPosNPC(nPos.Map, nPos.X, nPos.Y + 1, 0) Then
            'Mientras no halla bot.
            If Not .BotIndex <> 0 Then
               ia_HeadingToMolestNpc = eHeading.SOUTH
            End If
         End If
         
         'Pos legal hacia la izquierda?
         If LegalPosNPC(nPos.Map, nPos.X - 1, nPos.Y, 0) Then
            'Mientras no halla bot.
            If Not .BotIndex <> 0 Then
               ia_HeadingToMolestNpc = eHeading.WEST
            End If
         End If
         
         'Pos legal hacia la derecha?
         If LegalPosNPC(nPos.Map, nPos.X + 1, nPos.Y, 0) Then
            'Mientras no halla bot.
            If Not .BotIndex <> 0 Then
               ia_HeadingToMolestNpc = eHeading.EAST
            End If
         End If
         
    End With
 
End Function
 
Function ia_Objetos(ByVal BotIndex As Byte) As WorldPos
 
    ' @designer     :  maTih.-
    ' @date         :  2012/02/01
    ' @note         :  Busca objetos valiosos en el area.
     
    Dim loopX   As Long
    Dim loopY   As Long
    Dim BotPos  As WorldPos
     
    BotPos = IA_Bot(BotIndex).Pos
     
    '********************************
     
    'borro esto ya que no libero esta parte : p
     
    '********************************
     
    ia_Objetos.Map = 0
 
End Function
 
Function ia_SlotInventario(ByVal BotIndex As Byte) As Byte
 
    ' @designer     :  maTih.-
    ' @date         :  2012/02/01
    ' @note         :  Busca un slot libre.
     
    Dim loopX   As Long
     
    For loopX = 1 To IA_SLOTS
        With IA_Bot(BotIndex).Inv(loopX)
             'No hay objeto.
             If Not .ObjIndex <> 0 Then
                ia_SlotInventario = CByte(loopX)
                Exit Function
             End If
        End With
    Next loopX
     
    ia_SlotInventario = 0
 
End Function
 
Sub ia_ActionViajante(ByVal BotIndex As Byte)
 
    ' @designer     :  maTih.-
    ' @date         :  2012/02/01
    ' @note         :  Acciones de los bots que viajan hacia mapas.
     
    Dim RutaDir     As eHeading
    Dim molestNpc   As Integer
    Dim ObjetoPos   As WorldPos
     
    With IA_Bot(BotIndex)
     
         'Está paralizado?
         If .Paralizado Then
            'Puede tirar hechizos.
            If .Intervalos.SpellCount = 0 Then
               'se remueve
               ia_SendToBotArea BotIndex, PrepareMessageChatOverHead("AN HOAX VORP", .Char.CharIndex, vbCyan)
               .Paralizado = False
               .Intervalos.SpellCount = (IA_SINT / 30)
            End If
         End If
            
         'Se puede mover?
         If Not .Intervalos.MoveCharCount = 0 Then Exit Sub
            
         .Intervalos.MoveCharCount = (IA_MOVINT / 50)
         
         'Tiene una ruta?
         RutaDir = ia_HayRuta(.Pos)
        
         'Ve un objeto valioso?
         ObjetoPos = ia_Objetos(BotIndex)
         
         If ObjetoPos.Map <> 0 Then
            'Lo va a buscar, pero antes , setea su vieja pos.
            If Not .UltimaIdaObjeto Then
                .ViajanteAntes = .Pos
            End If
            
            ia_SearchPath BotIndex, ObjetoPos, RutaDir
            .UltimaIdaObjeto = True
         End If
         
         'No hay ruta?
         If Not RutaDir <> 0 Then
            'habia atacado un usuario? si es así volvemos a la pos.
            ia_SearchPath BotIndex, .ViajanteAntes, RutaDir
         End If
         
         If RutaDir <> 0 Then
            
            'Hacia donde mueve?
            Select Case RutaDir
                    
                   Case eHeading.NORTH      '<Mueve norte.
                        'Hay npc en su camino?
                        molestNpc = MapData(.Pos.Map, .Pos.X, .Pos.Y - 1).NpcIndex
                        
                        #If Barcos <> 0 Then
                            If molestNpc <> 0 Then
                                ia_SendToBotArea BotIndex, PrepareMessageChatOverHead("¡Maldita criatura, obstruyes mi paso!", .Char.CharIndex, vbWhite)
                                Call MoveNPCChar(molestNpc, ia_HeadingToMolestNpc(molestNpc))
                            End If
                        #End If
                        
                   Case eHeading.SOUTH      '<Mueve sur.
                        'Hay npc en su camino?
                        molestNpc = MapData(.Pos.Map, .Pos.X, .Pos.Y + 1).NpcIndex
                        
                        If molestNpc <> 0 Then
                           ia_SendToBotArea BotIndex, PrepareMessageChatOverHead("¡Maldita criatura, obstruyes mi paso!", .Char.CharIndex, vbWhite)
                           'muevo el npc
                           Call MoveNPCChar(molestNpc, ia_HeadingToMolestNpc(molestNpc))
                        End If
                           
                   Case eHeading.EAST       '<Mueve este.
                        'Hay npc en su camino?
                        molestNpc = MapData(.Pos.Map, .Pos.X + 1, .Pos.Y).NpcIndex
                        
                        If molestNpc <> 0 Then
                           ia_SendToBotArea BotIndex, PrepareMessageChatOverHead("¡Maldita criatura, obstruyes mi paso!", .Char.CharIndex, vbWhite)
                           'muevo el npc
                           Call MoveNPCChar(molestNpc, ia_HeadingToMolestNpc(molestNpc))
                        End If
                        
                   Case eHeading.WEST       '<Mueve oeste.
                        'Hay npc en su camino?
                        molestNpc = MapData(.Pos.Map, .Pos.X - 1, .Pos.Y).NpcIndex
                        
                        If molestNpc <> 0 Then
                           ia_SendToBotArea BotIndex, PrepareMessageChatOverHead("¡Maldita criatura, obstruyes mi paso!", .Char.CharIndex, vbWhite)
                           'Call MoveNPCChar(molestNpc, ia_HeadingToMolestNpc(molestNpc))
                        End If
            End Select
            
            'Move:p
            ia_MoveViajante BotIndex, RutaDir
            'Set el heading.
            .Char.heading = RutaDir
         End If
         
     
         
         'Encontramos una salida? - translados.
         If MapData(.Pos.Map, .Pos.X, .Pos.Y).TileExit.Map <> 0 Then
            'Mapa válido?
            If MapaValido(MapData(.Pos.Map, .Pos.X, .Pos.Y).TileExit.Map) Then
                'Asignamos nuevas posiciones, borramos el char anterior.
                ia_SendToBotArea BotIndex, PrepareMessageCharacterRemove(.Char.CharIndex)
                'Pos del npc.
                .Pos.Map = MapData(.Pos.Map, .Pos.X, .Pos.Y).TileExit.Map
                
                'Por si no tiene heading.
                If Not .Char.heading <> 0 Then .Char.heading = eHeading.SOUTH
                
                'Nueva X?
                If MapData(.Pos.Map, .Pos.X, .Pos.Y).TileExit.X <> 0 Then
                    .Pos.X = MapData(.Pos.Map, .Pos.X, .Pos.Y).TileExit.X
                End If
                
                'Nueva Y?
                If MapData(.Pos.Map, .Pos.X, .Pos.Y).TileExit.Y <> 0 Then
                    .Pos.Y = MapData(.Pos.Map, .Pos.X, .Pos.Y).TileExit.Y
                End If
                
                 MapData(.Pos.Map, .Pos.X, .Pos.Y).BotIndex = BotIndex
                'Creamos.
                
                Dim tmp_Color   As eNickColor
                
                'preparo el color del nick
                If .EsCriminal Then
                   tmp_Color = eNickColor.ieCriminal
                Else
                   tmp_Color = eNickColor.ieCiudadano
                End If
                
                ia_SendToBotArea BotIndex, PrepareMessageCharacterCreate(.Char.Body, .Char.Head, .Char.heading, .Char.CharIndex, .Pos.X, .Pos.Y, .Char.WeaponAnim, .Char.ShieldAnim, 0, 0, .Char.CascoAnim, .Name, tmp_Color, 0, 0)
            End If
         End If
         
    End With
 
End Sub
 
Function ia_HayRuta(ByRef InPos As WorldPos) As eHeading
     
    ' @designer     :  maTih.-
    ' @date         :  2012/02/01
    ' @note         :  Devuelve la dircción de la ruta en una pos.
     
    With MapData(InPos.Map, InPos.X, InPos.Y)
         
         'ia_HayRuta = .Rutas(1)
         
    End With
 
End Function
 
Public Sub CargarBOTs()
'*****************************************
'Autor: Lorwik
'Fecha: 01/05/2020
'Descripción: Carga los bots desde un Dat
'*****************************************

    Dim BOT         As Integer
    Dim LoopC       As Byte
    Dim ln          As String
    
    Dim Leer        As clsIniManager

    Set Leer = New clsIniManager
    
    Call Leer.Initialize(DatPath & "BOTs.dat")
    
    'obtiene el numero de BOTs
    MAX_BOTS = val(Leer.GetValue("INIT", "NumBOTs"))
    
    ReDim Preserve IA_Bot(1 To MAX_BOTS) As BOT
    
    'Reservamos los char
    IA_CHAR = (MAXCHARS - MAX_BOTS)
    
    For BOT = 1 To MAX_BOTS
        With IA_Bot(BOT)
    
            .Name = Leer.GetValue("BOT" & BOT, "Name")
            .Level = val(Leer.GetValue("BOT" & BOT, "Level"))
            .Raza = val(Leer.GetValue("BOT" & BOT, "Raza"))
            .Genero = val(Leer.GetValue("BOT" & BOT, "Genero"))
            .Char.Head = val(Leer.GetValue("BOT" & BOT, "Head"))
            .clase = val(Leer.GetValue("BOT" & BOT, "Clase"))
            .EsCriminal = CBool(Leer.GetValue("BOT" & BOT, "EsCriminal"))
            .minVida = val(Leer.GetValue("BOT" & BOT, "MinVida"))
            .maxVida = val(Leer.GetValue("BOT" & BOT, "MaxVida"))
            .minMana = val(Leer.GetValue("BOT" & BOT, "MinMana"))
            .maxMana = val(Leer.GetValue("BOT" & BOT, "MaxMana"))
            .Viajante = CBool(Leer.GetValue("BOT" & BOT, "Viajante"))
            
            .NroItems = val(Leer.GetValue("BOT" & BOT, "NROITEMS"))
            'Si el BOT Tiene inventario lo cargamos
            If .NroItems > 0 Then
                For LoopC = 1 To .NroItems
                    'No podemos superar la capacidad del inventario
                    If LoopC <= IA_SLOTS Then
                        ln = Leer.GetValue("BOT" & BOT, "Drop" & LoopC)
                        .Inv(LoopC).ObjIndex = val(ReadField(1, ln, 45))
                        .Inv(LoopC).Amount = val(ReadField(2, ln, 45))
                    End If
                Next LoopC
            End If
            
        End With
    Next BOT
    
    Set Leer = Nothing
End Sub

Public Sub CargarMensajesBOTS()
'*****************************************
'Autor: Lorwik
'Fecha: 01/05/2020
'Descripción: Carga los mensajes de los bots
'*****************************************
    Dim NumMsg      As Integer
    Dim LoopC       As Integer
    Dim Leer        As clsIniManager

    Set Leer = New clsIniManager
    
    Call Leer.Initialize(DatPath & "BOTsMsg.dat")
    
    NumMsg = val(Leer.GetValue("INIT", "NumMsg"))
    
    '¿Hay mensajes para cargar?
    If NumMsg = 0 Then Exit Sub
    
    For LoopC = 1 To NumMsg
        IA_Chats(LoopC) = Leer.GetValue("INIT", "Msg" & LoopC)
    Next LoopC
End Sub
 
Sub ia_SupportOthers(ByVal BotIndex As Byte, ByRef Supported As Boolean)
 
    ' @designer     :  maTih.-
    ' @date         :  2012/02/01
    ' @note         :  Un bot supportea otro.
     
    Dim botIndexToSupport   As Byte
    Dim supportAction       As eIASupportActions
     
    'Si no tiene intervalo..
    If IA_Bot(BotIndex).Intervalos.SpellCount <> 0 Then Exit Sub
     
    'Busca un bot a ayudar.
    botIndexToSupport = ia_GetSupportBot(BotIndex, supportAction)
     
    'No encontró, no supportea..
    If Not botIndexToSupport <> 0 Then Supported = False: Exit Sub
     
    'Que acción debe realizar?
    Select Case supportAction
     
           Case eIASupportActions.SCurar        '<Cura un compañero
                'Lanza graves.
                'Crea fx.
                ia_SendToBotArea botIndexToSupport, Protocol.PrepareMessageCreateFX(IA_Bot(botIndexToSupport).Char.CharIndex, Hechizos(5).FXgrh, Hechizos(5).loops)
                
                'Cartel.
                ia_SendToBotArea BotIndex, PrepareMessageChatOverHead("EN CORP SANCTIS", IA_Bot(BotIndex).Char.CharIndex, vbCyan)
                
                'Suma un random de vida.
                IA_Bot(botIndexToSupport).minVida = IA_Bot(botIndexToSupport).maxVida + RandomNumber(55, 77)
                
                'PARA QUE NO PASE LA VIDA MAXIMA
                If IA_Bot(botIndexToSupport).minVida > IA_Bot(botIndexToSupport).maxVida Then IA_Bot(botIndexToSupport).minVida = IA_Bot(botIndexToSupport).maxVida
           
                Supported = True
           
          Case eIASupportActions.SRemover       '<Remueve paralizis.
                'Crea el fx, el remo no tiene fx.
                'ia_sendtobotarea botindextosupport
                
                'Paralizis count.
                If IA_Bot(botIndexToSupport).Intervalos.ParalizisCount > 6 Then Exit Sub
                
                'Cartel
                ia_SendToBotArea BotIndex, PrepareMessageChatOverHead("AN HOAX VORP", IA_Bot(BotIndex).Char.CharIndex, vbCyan)
                
                'Saca el flag
                IA_Bot(botIndexToSupport).Paralizado = False
                
                Supported = True
                
    End Select
 
End Sub
 
Function ia_BotEnArea(ByVal BotIndex As Byte, ByVal otherBotIndex As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
     
    Dim BotIndexPos As WorldPos
     
    BotIndexPos = IA_Bot(BotIndex).Pos
     
    Dim loopX   As Long
    Dim loopY   As Long
     
    For loopY = BotIndexPos.Y - MinYBorder + 1 To BotIndexPos.Y + MinYBorder - 1
            For loopX = BotIndexPos.X - MinXBorder + 1 To BotIndexPos.X + MinXBorder - 1
                'hay un bot
                If MapData(BotIndexPos.Map, loopX, loopY).BotIndex = otherBotIndex Then
                    ia_BotEnArea = True
                    Exit Function
                End If
            
            Next loopX
    Next loopY
     
    ia_BotEnArea = False
 
End Function
 
Function ia_GetSupportBot(ByVal BotIndex As Byte, ByRef SAction As eIASupportActions) As Byte
 
    ' @designer     :  maTih.-
    ' @date         :  2012/02/01
    ' @note         :  Busca un bot a ayudar.
     
    Dim loopX   As Long
     
    For loopX = 1 To MAX_BOTS
        
        'Si no es mi BotIndex
        If loopX <> BotIndex Then
            
           'Está invocado?
           If IA_Bot(loopX).Invocado Then
              'Está en el area?
              If ia_BotEnArea(BotIndex, loopX) Then
                 'Está paralizado/tiene poca vida?
                 If IA_Bot(loopX).minVida <> IA_Bot(loopX).maxVida Or IA_Bot(loopX).Paralizado Then
                    'Encontrado.
                    ia_GetSupportBot = CByte(loopX)
                    'Devuelve la acción.
                    SAction = IIf(IA_Bot(loopX).minVida <> IA_Bot(loopX).maxVida, eIASupportActions.SCurar, eIASupportActions.SRemover)
                    Exit Function
                 End If
              End If
           End If
           
        End If
        
    Next loopX
 
ia_GetSupportBot = 0
End Function
 
Sub ia_Action(ByVal BotIndex As Byte)
 
On Error GoTo Errhandler        '< maTih XD
 
    ' @designer     :  maTih.-
    ' @date         :  2012/02/01
    ' @note         :  Acciones de los bots.
     
    Dim pIndex      As Integer
    Dim sRandom     As Integer
    Dim rMan        As Integer
    Dim FoundErr    As Boolean
    Dim moveHeading As eHeading
    Dim AyudoBot    As Boolean
     
    If EnPausa Then Exit Sub
     
    With IA_Bot(BotIndex)
     
        'Es un bot viajante?
        If .Viajante Then
              'Mientras no esté contra ningún pibe
              If Not .ViajanteUser <> 0 Then
                 ia_CheckInts BotIndex
                 ia_ActionViajante BotIndex
                 Exit Sub
              End If
        End If
        
        'si no lo ataco nadie  busca un target
        If (.ViajanteUser = 0) Then
            pIndex = ia_FindTarget(.Pos, .EsCriminal)
        Else
            pIndex = .ViajanteUser
        End If
        
        'No hay usuario.
        If pIndex <= 0 Then Exit Sub
     
        'Contadores de intervalo.
        ia_CheckInts BotIndex
       
        'EL bot boquea XD
        If Not .Intervalos.ChatCount <> 0 Then
           .Intervalos.ChatCount = (IA_TALKIN / 40)
            
           'Envia msj random
           ia_SendToBotArea BotIndex, PrepareMessageChatOverHead(IA_Chats(RandomNumber(1, 5)), .Char.CharIndex, -1)
           .Intervalos.SpellCount = (IA_SINT / 100)
        End If
        
        'Si se puede mover AND no está inmo se mueve al azar.
        If .Intervalos.MoveCharCount = 0 And .Paralizado = False Then
            
            'Tiene target?
            If pIndex <> 0 Then
               'busco un path.
               ia_SearchPath BotIndex, UserList(pIndex).Pos, moveHeading
            End If
            
            'Es clero?
            If Not .clase <> eIAClase.Clerigo Then
               'Si tiene la vida llena lo persigue.
               If .minVida = .maxVida Then
                  ia_MoveToHeading BotIndex, moveHeading, FoundErr
               Else
                'Si no , se mueve al azar.
                  ia_RandomMoveChar BotIndex, pIndex, FoundErr
               End If
             End If
                       
             'Es mago?
            If .clase = eIAClase.Mago Or .clase = eIAClase.Cazador Then
               'Si no tiene la vida llena se mueve al azar.
               If Not .minVida = .maxVida Then
                  ia_RandomMoveChar BotIndex, pIndex, FoundErr
               Else
                  'Tiene la vida llena, que fue el ultimo movimiento?
                  'Siguio la victima?
                  If .UltimoMovimiento = eIAMoviments.SeguirVictima Then
                     'Mueve random.
                     ia_RandomMoveChar BotIndex, pIndex, FoundErr
                     'Seteo.
                     .UltimoMovimiento = eIAMoviments.MoverRandom
                  Else
                     'Se movió al azar, sigue su victima.
                     ia_MoveToHeading BotIndex, moveHeading, FoundErr
                     'Seteo el nuevo flag.
                     .UltimoMovimiento = eIAMoviments.SeguirVictima
                 End If
            End If
           End If
           
           'se movio.
            If Not FoundErr Then
              'Se movió, guardo el BotIndex.
              MapData(.Pos.Map, .Pos.X, .Pos.Y).BotIndex = BotIndex
              
              'NEW--------
              'Checkeo si es una posición válida.
        
              'Actualizamos.
              ia_SendToBotArea BotIndex, PrepareMessageCharacterMove(.Char.CharIndex, .Pos.X, .Pos.Y)
              
              .Intervalos.MoveCharCount = (IA_MOVINT / 40)
            End If
            
        End If
       
       
        'STATS..
       
            'Prioriza la vida ante todo
           
            If .minVida < .maxVida Then
               
                'Checkeo el intervalo.
                If .Intervalos.UseItemCount > 0 Then Exit Sub
               
                'Recupera 20 cada 200 ms.
                .minVida = .minVida + 20
               
                If .minVida > .maxVida Then .minVida = .maxVida
               
                'Uso la poción, seteo el interval
                .Intervalos.UseItemCount = (IA_USEOBJ / 40)
               
                Exit Sub
            End If
           
            'Si tenia la vida llena usa azules.
           
            If .minMana < .maxMana Then
           
                'Checkeo el intervalo.
               
                If .Intervalos.UseItemCount = 0 Then
                
                    Dim recuperoMana    As Long
                    
                    'Recupera un % de la mana.
                    If .clase <> eIAClase.Mago Then
                        recuperoMana = Porcentaje(.maxMana, 5)
                    Else
                        recuperoMana = Porcentaje(.maxMana, 3)
                    End If
                    
                    'aumento el mana
                    .minMana = .minMana + recuperoMana
               
                    'controlo el limite
                    If .minMana > .maxMana Then .minMana = .maxMana
               
                'seteo el int
                .Intervalos.UseItemCount = (IA_USEOBJ / 40)
     
                End If
               
                'Hacer una constante después, con esto hacemos un random
                'Para que azulee y combee a la ves.
                If RandomNumber(1, 4) < 4 Then Exit Sub
            End If
       
        'Bueno si está acá es por que tenia la vida y mana llenas.
         
        'Es cazador??
        If .clase = eIAClase.Cazador Then
           'Intervalo permite?
           If Not .Intervalos.ArrowCount = 0 Then Exit Sub
           'Kza manqea XD - 25% de prob fallar
           If RandomNumber(1, 100) > 65 Then Exit Sub
           'Probabilidad de evadir.
           If Not RandomNumber(1, 100) <= MaximoInt(10, MinimoInt(90, 50 + ((220 - PoderEvasion(pIndex)) * 0.4))) Then
              'Atacó y falló!!
              Call WriteConsoleMsg(pIndex, .Name & " Te lanzó un flechazo pero falló!", FontTypeNames.FONTTYPE_FIGHT)
              'setea intervalo
              .Intervalos.ArrowCount = (IA_PROINT / 25)
              Exit Sub
           End If
           
           Dim ArrowDamage  As Integer  '<DañoBase.
           Dim ArmourIndex  As Integer  '<ArmaduraObjIndex
           Dim HelmetIndex  As Integer  '<CascoObjIndex
           
           ArrowDamage = RandomNumber(185, 225)
           
           'Restamos si tiene armadura.
           ArmourIndex = UserList(pIndex).Invent.ArmourEqpObjIndex
           HelmetIndex = UserList(pIndex).Invent.CascoEqpObjIndex
           
           'Pega en cabeza?
           If RandomNumber(1, 6) = 6 Then
              'Absorve.
              If HelmetIndex <> 0 Then
                 ArrowDamage = ArrowDamage - RandomNumber(ObjData(HelmetIndex).MinDef, ObjData(HelmetIndex).MaxDef)
              End If
           Else
              'Armadura absorce.
              If ArmourIndex <> 0 Then
                 ArrowDamage = ArrowDamage - RandomNumber(ObjData(ArmourIndex).MinDef, ObjData(ArmourIndex).MaxDef)
              End If
           End If
           
           'crea fx.
           'SendData SendTarget.ToPCArea, pIndex, mod_DunkanProtocol.Send_CreateArrow(.Char.CharIndex, UserList(pIndex).Char.CharIndex, ObjData(553).GrhIndex)
           
           'crea daño
           'Call mod_DunkanGeneral.Enviar_DañoAUsuario(pIndex, ArrowDamage)
           
           'Sacude un flechazo.
           UserList(pIndex).Stats.MinHp = UserList(pIndex).Stats.MinHp - ArrowDamage
           
           Call WriteConsoleMsg(pIndex, .Name & " Te ha pegado un flechazo por " & ArrowDamage, FontTypeNames.FONTTYPE_FIGHT)
           
           'Muere?
           If UserList(pIndex).Stats.MinHp <= 0 Then
              UserDie pIndex
              Call WriteConsoleMsg(pIndex, .Name & " Te ha matado!", FontTypeNames.FONTTYPE_FIGHT)
           End If
            
           'Intervalo
           .Intervalos.ArrowCount = (IA_PROINT / 20)
            
           'client update
           WriteUpdateHP pIndex
           Exit Sub
        End If
        
        'Puede castear?
        'Si el usuario no tiene la vida llena ataca
        Dim tmpHP   As Long
        
        tmpHP = (UserList(pIndex).Stats.MinHp)
        
        'obtengo el % de vida del user
        tmpHP = (tmpHP * 100) / (UserList(pIndex).Stats.MaxHp)
       
        If .Intervalos.SpellCount = 0 Then
        
        'Es clérigo y puede pegar??
        If (.clase = eIAClase.Clerigo) And .Intervalos.HitCount = 0 And Not .UltimaAccion = eIAactions.ePegar Then
           'Está al alcance de la víctima para un gole meele?
           Dim newBotHeading   As eHeading
           
           If ia_PuedeMeele(.Pos, UserList(pIndex).Pos, newBotHeading) Then
                'Acierta el golpe?
                If ia_AciertaGolpe(pIndex) Then
                   'Antes que nada cambiamos el heading, si es válido.
                   If newBotHeading <> 0 And newBotHeading <> .Char.heading Then
                      'ia_SendToBotArea botIndex, mod_DunkanProtocol.Send_ChangeHeadingChar(.Char.CharIndex, newBotHeading)
                   End If
                   
                   'Calcula el golpe
                   Dim GolpeVal     As Integer
                   GolpeVal = ia_CalcularGolpe(pIndex)
                   
                   'Resta hp.
                   UserList(pIndex).Stats.MinHp = UserList(pIndex).Stats.MinHp - GolpeVal
                   
                   'crea el fx de la sangre.
                   SendData SendTarget.ToPCArea, pIndex, PrepareMessageCreateFX(UserList(pIndex).Char.CharIndex, FXSANGRE, 5)
                   
                   'Avisa.
                   Call WriteConsoleMsg(pIndex, .Name & " Te ha pegado por " & CStr(GolpeVal) & ".", FontTypeNames.FONTTYPE_FIGHT)
                   
                   'Setea flag.
                   .UltimaAccion = eIAactions.ePegar
                   
                   'Muere?
                   If UserList(pIndex).Stats.MinHp <= 0 Then
                      Call UserDie(pIndex)
                   End If
                   
                   'update hp.
                   WriteUpdateHP pIndex
                   
                   'Intervalo de golpe.
                   .Intervalos.HitCount = (IA_HITINT / 40)
                   'Intervalo de hechizo.
                   .Intervalos.SpellCount = (IA_SINT / 40)
                   'Intervalo de golpe+pociones.
                   .Intervalos.UseItemCount = (IA_USEOBJ / 60)
                   Exit Sub
                End If
            End If
        End If
        
           'Feo, aunque digamos que solo hace apoca desc remo
           'Así que va a andar bien.
           
           'Si la mana es < a 300 [gasto del remo] no hacemos nada.
           
           If .minMana < 300 Then Exit Sub
           
           'Si está paralizado AND el usuario no tiene poka vida prioriza removerse.
           
            If .Paralizado And tmpHP > 60 Then
                
                'Intervalo de remo :@
                If .Intervalos.ParalizisCount <> 0 Then Exit Sub
                                'Palabras mágicas.

                'Palabras mágicas.
                
                ia_SendToBotArea BotIndex, PrepareMessagePalabrasMagicas(10, .Char.CharIndex)
               
                .Paralizado = False
               
                'Agrego esto por que si no tirarle inmo era al pedo
                'Seguia caminando practicamente :PP
               
                .Intervalos.ParalizisCount = (IA_SREMO / 10)
               
                'Se removió entonces salimos del sub y seteamos el intervalo
               
                .Intervalos.SpellCount = (IA_SINT / 40)
               
                Exit Sub
               
            End If
           
            'No está paralizado entonces castea un hechizo random.
           
            'chances de pegar
            If RandomNumber(1, 100) > IA_CASTEO Then Exit Sub
           
            sRandom = RandomNumber(1, IA_M_SPELL)
           
            'Ayuda otros bots si es que hay
            If NumInvocados <> 1 Then
               ia_SupportOthers BotIndex, AyudoBot
               
               'ayudo ya al bot?
               If AyudoBot Then
                  'SETEA INTERVALO
                  .Intervalos.SpellCount = (IA_SINT / 40)
                  Exit Sub
               End If
               
            End If
               
            'Si el usuario ya estaba paralizado AND el random es paralizar, entonces buscamos de nuevo
            If UserList(pIndex).flags.Paralizado = 1 And sRandom = 3 Then sRandom = RandomNumber(1, IA_M_SPELL - 1)
            
            'Si soy mago y el usuario es mago también no paraliza.
            If UserList(pIndex).clase = eClass.Mage And .clase = eIAClase.Mago Then sRandom = RandomNumber(1, IA_M_SPELL - 1)
            
            'Si el usuario tiene menos del 75% de vida juega al ataque.
            
            If tmpHP < 75 Then sRandom = RandomNumber(1, IA_M_SPELL - 1)
            
            'Si no llega con la mana del hechizo AND la del otro
            'tampoco entonces no hacemos nada
           
            If sRandom = 1 Then
               
                'Si no llega a la mana del spell 1 (descarga)
                'No hacemos nada ya que tampoco llega
                'al apocalipsis.
               
                rMan = Hechizos(IA_spell(1).spellIndex).ManaRequerido
               
                If rMan > .minMana Then Exit Sub
               
            ElseIf sRandom = 2 Then
           
                rMan = Hechizos(IA_spell(2).spellIndex).ManaRequerido
                   
                'Pero si es spell 2 (apoca) AND llegamos
                'con la mana para descarga, entonces
                'Seteamos sRandom como 1 y casteamos
                'descarga.
               
                If rMan > .minMana Then
                   
                    'Modifico la formula y hago un random
                    'Dado a que una ves que queda en -1000 de mana
                    'Nunca más tira apoca y castea puras descargas.
                   
                    If .minMana > 460 And RandomNumber(1, 100) < 30 Then
                        sRandom = 1
                    Else
                        Exit Sub
                    End If
                End If
           End If
           
            rMan = Hechizos(IA_spell(sRandom).spellIndex).ManaRequerido
           
            'Descontamos la maná y seteamos el intervalo.
            .minMana = .minMana - rMan
           
            'Set última action.
            .UltimaAccion = eIAactions.eMagia
            
            .Intervalos.SpellCount = (IA_SINT / 20) 'Se chekea cada 40 ms.
           
            'Creamos el fx y le descontamos la vida al usuario.
           
            ia_SendToBotArea BotIndex, Protocol.PrepareMessageCreateFX(UserList(pIndex).Char.CharIndex, Hechizos(IA_spell(sRandom).spellIndex).FXgrh, Hechizos(IA_spell(sRandom).spellIndex).loops)
           
            ia_SendToBotArea BotIndex, PrepareMessagePalabrasMagicas(IA_spell(sRandom).spellIndex, Npclist(BotIndex).Char.CharIndex)
           
            'Paralizar?
            If sRandom = 3 Then
               'Paralizado : P
               UserList(pIndex).flags.Paralizado = 1
               UserList(pIndex).Counters.Paralisis = IntervaloParalizado
               WriteParalizeOK pIndex
               WriteConsoleMsg pIndex, .Name & " Te ha paralizado", FontTypeNames.FONTTYPE_FIGHT
            End If
           
            'Random damage :D
           
            sRandom = RandomNumber(IA_spell(sRandom).DamageMin, IA_spell(sRandom).DamageMax)
           
            'Al daño le restamos , si el usuario tiene, defensa mágica.
            If UserList(pIndex).Invent.AnilloEqpObjIndex <> 0 Then
               sRandom = sRandom - RandomNumber(ObjData(UserList(pIndex).Invent.AnilloEqpObjIndex).DefensaMagicaMin, ObjData(UserList(pIndex).Invent.AnilloEqpObjIndex).DefensaMagicaMax)
            End If
            
            'NO numeros negativos.
            If sRandom < 0 Then sRandom = 0
           
            'Quitamos daño.
            UserList(pIndex).Stats.MinHp = UserList(pIndex).Stats.MinHp - sRandom
                
            If sRandom <> 0 Then
                'AVISO AL USUARIO DE ESTO
                Call WriteConsoleMsg(pIndex, .Name & " Te ha quitado " & sRandom & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
            End If
            
            'Check si muere.
            If UserList(pIndex).Stats.MinHp <= 0 Then
                 UserDie pIndex
                 
                'Era viajante y mató el usuario?, resteo el ui
                 If Not pIndex <> .ViajanteUser Then
                    .ViajanteUser = 0
                 End If
                 
                 'aviso que murio.
                 Call WriteConsoleMsg(pIndex, .Name & " Te ha matado!", FontTypeNames.FONTTYPE_FIGHT)
                 
            End If
           
            'Actualizamos el cliente.
           
            WriteUpdateHP pIndex
           
        End If
    End With
     
    Exit Sub
     
Errhandler:
 
End Sub
 
Sub ia_EnviarChar(ByVal UserIndex As Integer, ByVal BotIndex As Byte)
 
    ' @designer     :  maTih.-
    ' @date         :  2012/03/13
    ' @                Envia el char del bot a un usuario (sistema de areas!!)
     
        With IA_Bot(BotIndex).Char
                Dim tmp_Color   As eNickColor
                
                If IA_Bot(BotIndex).EsCriminal Then
                   tmp_Color = eNickColor.ieCriminal
                Else
                   tmp_Color = eNickColor.ieCiudadano
                End If
                
                Call Protocol.WriteCharacterCreate(UserIndex, .Body, .Head, eHeading.SOUTH, .CharIndex, IA_Bot(BotIndex).Pos.X, IA_Bot(BotIndex).Pos.Y, .WeaponAnim, .ShieldAnim, 0, 0, .CascoAnim, IA_Bot(BotIndex).Name, tmp_Color, 0)
        End With
 
End Sub
 
Sub ia_UserDamage(ByVal spell As Byte, ByVal BotIndex As Byte, ByVal UserIndex As Integer, Optional ByVal is_RuneArea As Boolean = False)
 
    ' @designer     :  maTih.-
    ' @date         :  2012/02/01
     
    Dim rMan     As Integer
    Dim Damage   As Integer
    Dim usedFont As FontTypeNames
     
    usedFont = FontTypeNames.FONTTYPE_FIGHT
     
    'Checkeo que el hechizo no sea 0.
    If Not spell <> 0 Then Exit Sub
     
    With UserList(UserIndex)
     
        rMan = Hechizos(spell).ManaRequerido
       
        'Llega con la mana?
        If rMan > .Stats.MinMAN Then
            WriteConsoleMsg UserIndex, "No tienes suficiente mana!", usedFont
            Exit Sub
        End If
        
        'Skills?
        If Hechizos(spell).MinSkill > .Stats.UserSkills(eSkill.Magia) Then
           WriteConsoleMsg UserIndex, "No tienes suficientes skills en magia", usedFont
           Exit Sub
        End If
       
        'Soy ciudadano y el target es un bot viajante?
        If Not criminal(UserIndex) And IA_Bot(BotIndex).Viajante And .flags.Seguro Then
            WriteConsoleMsg UserIndex, "Para atacar bots viajantes debes desactivar el seguro", usedFont
            Exit Sub
        End If
        
        If Hechizos(spell).Inmoviliza Or Hechizos(spell).Paraliza Then
           
            'Le pongo el flag en verdadero.
            IA_Bot(BotIndex).Paralizado = True
           
            'Mensaje informando.
            WriteConsoleMsg UserIndex, "Has paralizado ah " & IA_Bot(BotIndex).Name, usedFont
            
            'Creo la animacion sobre el char.
            ia_SendToBotArea BotIndex, Protocol.PrepareMessageCreateFX(IA_Bot(BotIndex).Char.CharIndex, Hechizos(spell).FXgrh, Hechizos(spell).loops)
            
            'SpellWorlds.
            DecirPalabrasMagicas spell, UserIndex
           
            'Quito mana y energia
            .Stats.MinMAN = .Stats.MinMAN - rMan
           
            'le doy intervalo
           
            IA_Bot(BotIndex).Intervalos.ParalizisCount = (IA_SREMO / 10)
           
            WriteUpdateMana UserIndex
           
            Exit Sub
        End If
       
        'Era un Viajante
       
        Damage = RandomNumber(Hechizos(spell).MinHp, Hechizos(spell).MaxHp)
        Damage = Damage + Porcentaje(Damage, 3 * .Stats.ELV)
       
       If Not Damage <> 0 Then Exit Sub
       
       If IA_Bot(BotIndex).Viajante Then
            Dim eraPK   As Boolean
            
            If Not IA_Bot(BotIndex).ViajanteAntes.Map Then
                IA_Bot(BotIndex).ViajanteAntes = IA_Bot(BotIndex).Pos
            End If
            
            'No era criminal.
            eraPK = criminal(UserIndex)
        
            'No era criminal y atacó un viajante, es criminal.
            If Not eraPK Then VolverCriminal UserIndex
        
            'Ahora el bot se enojó viejo..
            IA_Bot(BotIndex).ViajanteUser = UserIndex
        
            'UserList(UserIndex).AtacoViajante = BotIndex
        
            WriteConsoleMsg UserIndex, "Has atacado un viajante!! ahora eres un criminal, y además el viajante te atacará!", usedFont
       End If
       
        'Quitamos vida
        If Hechizos(spell).StaffAffected Then
           If UserList(UserIndex).clase = eClass.Mage Then
              If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                 Damage = (Damage * (ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).StaffDamageBonus + 70)) / 100
              Else
                 Damage = Damage * 0.7 'Baja damage a 70% del original
              End If
            End If
         End If
            
         If UserList(UserIndex).Invent.AnilloEqpObjIndex = LAUDELFICO Or UserList(UserIndex).Invent.AnilloEqpObjIndex = FLAUTAELFICA Then
            Damage = Damage * 1.04  'laud magico de los bardos
         End If
        
        IA_Bot(BotIndex).minVida = IA_Bot(BotIndex).minVida - Damage
        
        'No está paralizado.
        If Not IA_Bot(BotIndex).Paralizado Then
            'Le pegaron, se cagó todo y se mueve random.
            Dim keepMoving  As Boolean
        
            ia_RandomMoveChar BotIndex, UserIndex, keepMoving
        
            'No hubo error, por ende se movió.
            If Not keepMoving Then
               'Guardo la nueva pos.
               MapData(IA_Bot(BotIndex).Pos.Map, IA_Bot(BotIndex).Pos.X, IA_Bot(BotIndex).Pos.Y).BotIndex = BotIndex
           
               'Actualizo el area del bot.
               ia_SendToBotArea BotIndex, PrepareMessageCharacterMove(IA_Bot(BotIndex).Char.CharIndex, IA_Bot(BotIndex).Pos.X, IA_Bot(BotIndex).Pos.Y)
           
               'Intervalo de caminata.
               IA_Bot(BotIndex).Intervalos.MoveCharCount = (IA_MOVINT / 40)
            End If
            
        End If
        
        'Aviso al usuario.
        WriteConsoleMsg UserIndex, "Le has quitado " & CStr(Damage) & " puntos de vida a " & IA_Bot(BotIndex).Name, usedFont
       
        'Tiro las spell worlds
        DecirPalabrasMagicas spell, UserIndex
       
        'Creo el fx.
        ia_SendToBotArea BotIndex, Protocol.PrepareMessageCreateFX(IA_Bot(BotIndex).Char.CharIndex, Hechizos(spell).FXgrh, Hechizos(spell).loops)
       
        'saco mana y energia y actualizo el cliente
        .Stats.MinMAN = .Stats.MinMAN - rMan
           
        WriteUpdateMana UserIndex
       
        If IA_Bot(BotIndex).minVida <= 0 Then
            'Murió?
            ia_EraseChar BotIndex, True
            WriteConsoleMsg UserIndex, "Has matado ah " & IA_Bot(BotIndex).Name & ".", usedFont
        End If
       
    End With
 
End Sub
 
Sub ia_DamageHit(ByVal BotIndex As Byte, ByVal UserIndex As Integer)
 
    ' @designer     :  maTih.-
    ' @date         :  2012/02/01
     
    Dim nDamage      As Integer
     
    'Calculo el daño.
    nDamage = CalcularDano(UserIndex)
     
    'Resto la defensa del bot.
    nDamage = nDamage - (RandomNumber(IA_MINDEF, IA_MAXDEF))
     
    'Aviso al usuario.
    WriteConsoleMsg UserIndex, "Le has pegado a " & IA_Bot(BotIndex).Name & " por " & nDamage, FontTypeNames.FONTTYPE_FIGHT
     
    'Creo daño :)
    'ia_SendToBotArea BotIndex, mod_DunkanProtocol.Send_CreateDamage(ia_Bot(BotIndex).Pos.X, ia_Bot(BotIndex).Pos.Y, nDamage)
     
    'Resto vida.
    IA_Bot(BotIndex).minVida = IA_Bot(BotIndex).minVida - nDamage
     
    'seteo el flag.
    'UserList(UserIndex).AtacoViajante = BotIndex
     
    'Murio?
    If IA_Bot(BotIndex).minVida <= 0 Then
     
        'Era viajante?
        If IA_Bot(BotIndex).Viajante Then
           'Reset el flag.
           'UserList(UserIndex).AtacoViajante = 0
        End If
        
        ia_EraseChar BotIndex, True
        
    End If
 
End Sub
 
Sub ia_SendToBotArea(ByVal BotIndex As Byte, ByVal PackData As String)
 
    ' @designer     :  maTih.-
    ' @date         :  2012/03/13
    ' @                Envia paquetes al area de un bot.
     
    'Nueva versión del sub, más simple y diría que más práctica : P
     
    With IA_Bot(BotIndex)
        'd3 ao, borro esto : p
        
        'con esto tenemos algo simple, cuando mandamos el send
        'tobotarea, nos devuelve un array con los ui y el ping de cada
        'uno, y flush_ping tiene el promedio :), despues solo nos
        'queda comprobar si el usuario puede flushbuffear los datos
        'y enviamos, sacrificamos memoria pero ganamos MUCHA conexión.
        
        'Dim flush_Ping      As Integer
        'Dim arr_PingUsers() As Integer
        
        'Call modSendData.SendToAreaByPos(.Pos.map, .Pos.X, .Pos.Y, PackData, .GrupoID, flush_Ping)
        
        'Do While flush_Ping <> 0
        '    If can_Update_Ping(arr_PingUsers(flush_Ping)) Then
        '       Call flusH_buffer_to_base_Ping(arr_PingUsers(flush_Ping), flush_Ping, .GrupoID)
        '    End If
            
        '    flush_Ping = flush_Ping - 1
            
        'Loop
        
        Call modSendData.SendToAreaByPos(.Pos.Map, .Pos.X, .Pos.Y, PackData)
    End With
 
End Sub
 
Sub IA_TirarInventario(ByVal BotIndex As Byte)
'*************************************
'Autor: Lorwik
'Fecha 01/05/2020
'Descripcion: Si el BOT tiene inventario lo tiramos
'*************************************
    On Error GoTo Errhandler

    Dim i         As Byte

    Dim NuevaPos  As WorldPos

    Dim MiObj     As obj

    Dim ItemIndex As Integer

    Dim DropAgua  As Boolean
    
    With IA_Bot(BotIndex)
    
        '¿Tiene items en el inventario?
        If .NroItems = 0 Then Exit Sub

        For i = 1 To .NroItems
            ItemIndex = .Inv(i).ObjIndex

            If ItemIndex > 0 Then
                If ItemSeCae(ItemIndex) Then
                    NuevaPos.X = 0
                    NuevaPos.Y = 0
                    
                    'Creo el Obj
                    MiObj.Amount = .Inv(i).Amount
                    MiObj.ObjIndex = ItemIndex

                    DropAgua = True
                    
                    Call Tilelibre(.Pos, NuevaPos, MiObj, DropAgua, True)
                    
                    If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
                        MakeObj MiObj, NuevaPos.Map, NuevaPos.X, NuevaPos.Y
                    End If

                End If

            End If

        Next i

    End With
    
    Exit Sub
    
Errhandler:
    Call LogError("Error en TirarTodosLosItems. Error: " & Err.Number & " - " & Err.description)
 
End Sub
 
Sub ia_EraseChar(ByVal BotIndex As Byte, Optional ByVal killedbyUSER As Boolean = False)
 
    ' @designer     :  maTih.-
    ' @date         :  2012/02/01
    ' @note         :  Borra el char y los datos del bot.
     
    With IA_Bot(BotIndex)
        'Borro el char.
        ia_SendToBotArea BotIndex, PrepareMessageCharacterRemove(.Char.CharIndex)
        
        'Borro el botIndex
        MapData(.Pos.Map, .Pos.X, .Pos.Y).BotIndex = 0
        
        Dim dummyPos    As WorldPos
        
        .ViajanteAntes = dummyPos
        
        'Mató un usuario? pincha inventario
        If killedbyUSER Then
           IA_TirarInventario BotIndex
        End If
        
        'Reset char,
        With .Char
             .Body = 0
             .CascoAnim = 0
             .FX = 0
             .loops = 0
             .Head = 0
             .heading = 0
             .ShieldAnim = 0
             .WeaponAnim = 0
        End With
        
        'Reset STATS
        .minVida = 0
        .minMana = 0
        
        'Reset pos.
        With .Pos
             .Map = 0
             .X = 0
             .Y = 0
        End With
        
        'Reset flags.
        .Invocado = False
        .Paralizado = False
       
        'Reset intervalos.
        With .Intervalos
             .MoveCharCount = 0
             .SpellCount = 0
             .UseItemCount = 0
             .ParalizisCount = 0
        End With
        
        'Reset viajante flag.
        .Viajante = False
        .ViajanteUser = 0
        
        'Resta el contador
        NumInvocados = NumInvocados - 1
        
    End With
 
End Sub
 
Sub ia_CheckInts(ByVal BotIndex As Byte)
 
    ' @designer     :  maTih.-
    ' @date         :  2012/02/01
     
    With IA_Bot(BotIndex).Intervalos
         
        If .ArrowCount > 0 Then .ArrowCount = .ArrowCount - 1
        If .MoveCharCount > 0 Then .MoveCharCount = .MoveCharCount - 1
        If .SpellCount > 0 Then .SpellCount = .SpellCount - 1
        If .UseItemCount > 0 Then .UseItemCount = .UseItemCount - 1
        If .ParalizisCount > 0 Then .ParalizisCount = .ParalizisCount - 1
        If .HitCount > 0 Then .HitCount = .HitCount - 1
        If .ChatCount > 0 Then .ChatCount = .ChatCount - 1
        
    End With
 
End Sub
 
Function ia_FindTarget(Pos As WorldPos, Optional ByVal esPk As Boolean = False) As Integer
 
    ' @designer     :  maTih.-
    ' @date         :  2012/03/13
    ' @note         :  Busca alguien a quien pegar
     
    Dim loopX       As Long         '< Bucle del tileX.
    Dim loopY       As Long         '< Bucle del tileY.
    Dim tmpIndex    As Integer
     
    For loopY = Pos.Y - (MinYBorder + 1) To Pos.Y + (MinYBorder - 1)
        For loopX = Pos.X - (MinXBorder + 1) To Pos.X + (MinXBorder - 1)
            'Hay usuario?
            If MapData(Pos.Map, loopX, loopY).UserIndex > 0 Then
            
                '¿Es Admin o GM?
                If EsAdmin(UserList(MapData(Pos.Map, loopX, loopY).UserIndex).Name) Or EsGm(MapData(Pos.Map, loopX, loopY).UserIndex) Then Exit Function
                
                'No está muerto
                If UserList(MapData(Pos.Map, loopX, loopY).UserIndex).flags.Muerto = 0 Then
                
                      'Es ciuda el bot y el usuario?
                    If Not esPk Then
                        'el bot no es pk.
                        ia_FindTarget = MapData(Pos.Map, loopX, loopY).UserIndex
                    Else
                         tmpIndex = MapData(Pos.Map, loopX, loopY).UserIndex
                         If Not esPk And criminal(tmpIndex) Then
                             ia_FindTarget = tmpIndex
                         Else
                            If esPk And Not criminal(tmpIndex) Then
                               ia_FindTarget = tmpIndex
                            End If
                        End If
                    End If
                    Exit Function
                    
                End If
                
            End If
        Next loopX
    Next loopY
     
    ia_FindTarget = 0
End Function
 
Function IA_GetNextSlot() As Byte
 
    ' @ Devuelve un slot para bots.
     
    Dim loopX   As Long
     
    For loopX = 1 To MAX_BOTS
        If Not IA_Bot(loopX).Invocado Then
           IA_GetNextSlot = CByte(loopX)
           Exit Function
        End If
    Next loopX
     
    IA_GetNextSlot = 0
 
End Function
