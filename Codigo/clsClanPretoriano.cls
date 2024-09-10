VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsClanPretoriano"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************
' clsClanPretoriano - Handles the Praeorians NPCs.
'
' Designed (original version) by Mariano Barrou (El Oso)
' Redesigned by ZaMa
'**************************************************************

'**************************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'**************************************************************************
Option Explicit

'####################
' Pretorians Clan Type (Define wether it is default or custom clan)
Public Enum ePretorianType
    Default = 1
    Custom = 2
End Enum

'####################
' Pretorians Map Vars (Only used for the default map)
Private LeftSpawnPos  As WorldPos

Private RightSpawnPos As WorldPos

'####################
'Actions and spells

' King
Private Enum eKingSpells

    LittleHeal = 1
    RemoveParalisis = 2
    BlindAttack = 3
    StupidityAttack = 4
    HealPoison = 5
    HealQuemaduras = 6

End Enum

Private Enum eKingAction

    HealAllyParalisis = 1
    HealAllyPoison
    HealAlly
    ChaseTarget
    MassiveAttack
    RetreatAndHeal
    HealAllyQuemaduras

End Enum

' Healer
Private Enum eHealerAction

    HealAllyParalisis = 1
    ParalizePet
    ParalizeUser
    AttackUser
    HealAlly
    RescueFarAlly

End Enum

Private Enum eHealerSpells

    ParalizeUser = 1
    RemoveParalisis = 2
    ParalizeNpc = 3
    Heal = 4
    Tormenta = 5

End Enum

' SpellCaster
Private Enum eSpellCasterAction

    RemoveInvi = 1
    Attack

End Enum

Private Enum eSpellCasterSpells

    Apocalipsis = 1
    RemoInvi = 2

End Enum

' Thief
Private Enum eThiefAction

    Attack = 1
    Steal
    None

End Enum

Private Enum eThiefSpells

    Arrow = 1
    Paralisis = 2

End Enum

' Indicates wether the clan is active or not
Private ClanActive                   As Boolean

' Max distance for the thief to steal
Private Const THIEF_STEAL_DISTANCE   As Integer = 2

' Countdown to break the SpellCaster's wand
Private WandBreakCounter             As Integer

' Number from wich starts countdown
Private Const MAX_WAND_BREAK_VALUE   As Integer = 6

' Min HP needed for SpellCaster to start countdown
Private Const START_COUNTDOWN_MIN_HP As Integer = 750

' Time in which the king hasn't done anything
Private KingIdleTime                 As Long

' Time needed for respawning an ally (Aprox 5 min)
Private Const ALLY_RESPAWN_TIME      As Long = 2000

' Max distance to follow a user, or separate from team
Private Const MAX_DISTANCE           As Integer = 14

' Number of total clan members
Private Const NRO_PRETORIANOS        As Integer = 9

' Indicates clan map and king position
Private CenterPos                    As WorldPos

' Indicates if the clan is spawned after being defeated
Private RespawnClan                  As Boolean

' The number of clan members still alive
Private ClanMembersAlive             As Integer

Private Type tPretorianData

    NPCIndex As Integer
    NPCAI As ePretorianAI

End Type

' Indice del clan pretoriano
Private ClanIndex     As Integer

' Contiene los index de los miembros del clan.
Private Pretorianos() As tPretorianData

Public Function SpawnClan(ByVal Mapa As Integer, _
                          ByVal X As Integer, _
                          ByVal Y As Integer, _
                          ByVal PretoClanIndex As Integer, _
                          Optional ByVal Respawning As Boolean = False) As Boolean

    '********************************************************
    'Author: EL OSO
    'Inicializa el clan Pretoriano.
    'Last Modify Date: 21/09/2010
    '22/06/2006: (Nacho) - Seteamos cuantos NPCs creamos
    '21/09/2010: ZaMa - Reescribi la funcion, ahora se le pasa las coordenadas de respawn
    '********************************************************
    On Error GoTo ErrHandler

    Dim nPos       As WorldPos

    Dim PretoIndex As Integer
    
    ' Initial pos
    CenterPos.Map = Mapa
    CenterPos.X = X
    CenterPos.Y = Y

    ReDim Pretorianos(1 To NRO_PRETORIANOS) As tPretorianData
    
    ClanIndex = PretoClanIndex
    
    ' Check valid spawn place only if it's the first time
    If Not Respawning Then

        ' The clan only respawns on pretorian map
        If Mapa = MAPA_PRETORIANO Then
            RespawnClan = True
            Call InitializePretoriansVars
            
            ' Default coordinates, override admin choise
            CenterPos = LeftSpawnPos
            
            ' Check if it's a valid area to respawn
        Else

            If Not IsValidSpawnArea() Then Exit Function

        End If

    End If

    nPos.Map = Mapa
    nPos.X = X
    nPos.Y = Y

    ' Spawn clan members
    ClanMembersAlive = NRO_PRETORIANOS
    
    ' Spawn King
    PretoIndex = PretoIndex + 1
    Call SpawnPretorian(nPos, ePretorianAI.King, PretoIndex, True)
    
    ' Spawn 2 Healers
    PretoIndex = PretoIndex + 1
    nPos.X = nPos.X + 3
    Call SpawnPretorian(nPos, ePretorianAI.Healer, PretoIndex)
    
    PretoIndex = PretoIndex + 1
    nPos.X = nPos.X - 6
    Call SpawnPretorian(nPos, ePretorianAI.Healer, PretoIndex)
    
    ' Spawn 3 Sword Masters
    PretoIndex = PretoIndex + 1
    nPos.Y = nPos.Y + 3
    Call SpawnPretorian(nPos, ePretorianAI.SwordMaster, PretoIndex)
    
    PretoIndex = PretoIndex + 1
    nPos.X = nPos.X + 3
    Call SpawnPretorian(nPos, ePretorianAI.SwordMaster, PretoIndex)
    
    PretoIndex = PretoIndex + 1
    nPos.X = nPos.X + 3
    Call SpawnPretorian(nPos, ePretorianAI.SwordMaster, PretoIndex)
    
    ' Sapwn 1 Shooter
    PretoIndex = PretoIndex + 1
    nPos.Y = nPos.Y - 6
    nPos.X = nPos.X - 1
    Call SpawnPretorian(nPos, ePretorianAI.Shooter, PretoIndex)
    
    ' Spawn 1 SpellCaster
    PretoIndex = PretoIndex + 1
    nPos.X = nPos.X - 4
    Call SpawnPretorian(nPos, ePretorianAI.SpellCaster, PretoIndex)
    
    ' Spawn 1 Thief
    PretoIndex = PretoIndex + 1
    Call SpawnPretorian(nPos, ePretorianAI.Thief, PretoIndex)

    ' Reset wand break counter
    WandBreakCounter = MAX_WAND_BREAK_VALUE
    
    ' Now is active
    ClanActive = True
    
    ' Clan creation success
    SpawnClan = True

    Exit Function

ErrHandler:
    Call LogError("Error en CrearClanPretoriano. Error: " & Err.Number & " - " & Err.description)

End Function

Private Function IsValidSpawnArea() As Boolean
    '********************************************************
    'Author: ZaMa
    'Checks if it's a clean area to respawn the clan.
    'Required conditions:
    '-> A 12x12 square free of blocks and in map bounds.
    '-> Must be a non-water terrain
    'Last Modify Date: 21/09/2010
    '********************************************************
    
    Dim loopX As Long

    Dim loopY As Long
    
    For loopX = CenterPos.X - 6 To CenterPos.X + 6
        For loopY = CenterPos.Y - 6 To CenterPos.Y + 6
            
            If Not InMapBounds(CenterPos.Map, loopX, loopY) Then Exit Function
            
            If HayAgua(CenterPos.Map, loopX, loopY) Then Exit Function
            
        Next loopY
    Next loopX
    
    IsValidSpawnArea = True
    
End Function

Private Sub InitializePretoriansVars()
    '********************************************************
    'Author: ZaMa
    'Initialize the variables/const used in pretorian map
    'Last Modify Date: 21/09/2010
    '********************************************************

    LeftSpawnPos.Map = MAPA_PRETORIANO
    LeftSpawnPos.X = 830
    LeftSpawnPos.Y = 425
    
    RightSpawnPos.Map = MAPA_PRETORIANO
    RightSpawnPos.X = 867
    RightSpawnPos.Y = 425
    
End Sub

Private Sub SpawnPretorian(ByRef SpawnPos As WorldPos, _
                           ByVal PretorianAI As Byte, _
                           ByVal PretoIndex As Integer, _
                           Optional ByVal IsKing As Boolean = False)
    '********************************************************
    'Author: ZaMa
    'Spawns the Pretorian in the given position.
    'Last Modify Date: 21/09/2010
    '********************************************************

    Dim FinalPos       As WorldPos

    Dim NPCIndex       As Integer

    Dim Head           As Integer
    
    Dim PretoDatNumber As Integer

    PretoDatNumber = RandomNumber(PretorianAIOffset(PretorianAI), PretorianAIOffset(PretorianAI + 1) - 1)
                            
    ' Odd numbers are tall races
    If (PretoDatNumber And 1) = 1 Then
        Head = RandomTallHead
    Else
        Head = RandomShortHead

    End If

    ' The King has to spawn always in the given position
    If IsKing Then
        
        ' Check if there's any npc in the king's position
        Dim OtherNpcIndex As Integer

        OtherNpcIndex = MapData(SpawnPos.Map, SpawnPos.X, SpawnPos.Y).NPCIndex
        
        ' Found one
        If OtherNpcIndex <> 0 Then
            ' Move it to the closest legal position
            Call ClosestLegalPos(SpawnPos, FinalPos)

            If FinalPos.X <> 0 And FinalPos.Y <> 0 Then
                Call SendData(SendTarget.ToNPCArea, OtherNpcIndex, PrepareMessageCharacterMove(Npclist(OtherNpcIndex).Char.CharIndex, FinalPos.X, FinalPos.Y))
                ' Update map and npc pos
                MapData(SpawnPos.Map, SpawnPos.X, SpawnPos.Y).NPCIndex = 0
                Npclist(OtherNpcIndex).Pos = FinalPos
                MapData(FinalPos.Map, FinalPos.X, FinalPos.Y).NPCIndex = OtherNpcIndex
            
                ' King's position is now empty
                FinalPos = SpawnPos
            
                ' Couldn't find a legal pos => Remove it
            Else
                Call QuitarNPC(OtherNpcIndex)

            End If

        Else
            FinalPos = SpawnPos

        End If
        
        ' Other clan member doesn't have to spawn in the exact position
    Else
        Call ClosestLegalPos(SpawnPos, FinalPos, False, True)

    End If
                            
    ' Create and Spawn npc
    NPCIndex = CrearNPC(PretorianDatNumbers(PretoDatNumber), FinalPos.Map, FinalPos, Head)
    
    ' Store Index
    Pretorianos(PretoIndex).NPCIndex = NPCIndex
    Pretorianos(PretoIndex).NPCAI = PretorianAI
    
    ' Asign ClanIndex
    Npclist(NPCIndex).ClanIndex = ClanIndex
    
End Sub

Private Function RandomTallHead()
    '********************************************************
    'Author: ZaMa
    'Gets a random tall race's Head.
    'Last Modify Date: 21/09/2010
    '********************************************************
    
    ' 3 tall races
    Dim Raza As Integer

    Raza = RandomNumber(1, 3)
        
    ' Male
    If (RandomNumber(0, 1) = 1) Then
        
        ' Human
        If Raza = 1 Then
            RandomTallHead = RandomNumber(HUMANO_H_PRIMER_CABEZA, HUMANO_H_ULTIMA_CABEZA)
            ' Elf
        ElseIf Raza = 2 Then
            RandomTallHead = RandomNumber(ELFO_H_PRIMER_CABEZA, ELFO_H_ULTIMA_CABEZA)
            ' Drow
        Else
            RandomTallHead = RandomNumber(DROW_H_PRIMER_CABEZA, DROW_H_ULTIMA_CABEZA)

        End If
        
        ' Female
    Else
        
        ' Human
        If Raza = 1 Then
            RandomTallHead = RandomNumber(HUMANO_M_PRIMER_CABEZA, HUMANO_M_ULTIMA_CABEZA)
            ' Elf
        ElseIf Raza = 2 Then
            RandomTallHead = RandomNumber(ELFO_M_PRIMER_CABEZA, ELFO_M_ULTIMA_CABEZA)
            ' Drow
        Else
            RandomTallHead = RandomNumber(DROW_M_PRIMER_CABEZA, DROW_M_ULTIMA_CABEZA)

        End If
        
    End If
    
End Function

Private Function RandomShortHead()
    '********************************************************
    'Author: ZaMa
    'Gets a random short race's Head.
    'Last Modify Date: 21/09/2010
    '********************************************************

    ' 2 short races
    Dim Raza As Integer

    Raza = RandomNumber(1, 2)
        
    ' Male
    If (RandomNumber(0, 1) = 1) Then
        
        ' Dwarf
        If Raza = 1 Then
            RandomShortHead = RandomNumber(ENANO_H_PRIMER_CABEZA, ENANO_H_ULTIMA_CABEZA)
            ' Gnome
        Else
            RandomShortHead = RandomNumber(GNOMO_H_PRIMER_CABEZA, GNOMO_H_ULTIMA_CABEZA)

        End If
        
        ' Female
    Else
        
        ' Dwarf
        If Raza = 1 Then
            RandomShortHead = RandomNumber(ENANO_M_PRIMER_CABEZA, ENANO_M_ULTIMA_CABEZA)
            ' Gnome
        Else
            RandomShortHead = RandomNumber(GNOMO_M_PRIMER_CABEZA, GNOMO_M_ULTIMA_CABEZA)

        End If
        
    End If

End Function

Public Sub PerformPretorianAI(ByVal NPCIndex As Integer)
    '********************************************************
    'Author: ZaMa
    'Performs Pretorian's AI.
    'Last Modify Date: 21/09/2010
    '********************************************************
    
    Dim PretorianAI As Byte

    PretorianAI = GetPretorianAI(NPCIndex)
    
    Select Case PretorianAI
        
        Case ePretorianAI.King
            Call AI_King(NPCIndex)
            
        Case ePretorianAI.Healer
            Call AI_Healer(NPCIndex)
            
        Case ePretorianAI.SpellCaster
            Call AI_SpellCaster(NPCIndex)
            
        Case ePretorianAI.SwordMaster
            Call AI_SwordMaster(NPCIndex)
            
        Case ePretorianAI.Shooter
            Call AI_Shooter(NPCIndex)
            
        Case ePretorianAI.Thief
            Call AI_Thief(NPCIndex)
            
    End Select
    
End Sub

Private Function GetPretorianAI(ByVal NPCIndex As Integer) As Byte
    '********************************************************
    'Author: ZaMa
    'Returns the Pretorian's AI for the given Npc.
    'Last Modify Date: 21/09/2010
    '********************************************************
    
    Dim Counter As Long
    
    For Counter = 1 To NRO_PRETORIANOS

        If Pretorianos(Counter).NPCIndex = NPCIndex Then
            GetPretorianAI = Pretorianos(Counter).NPCAI
            Exit Function

        End If

    Next Counter
    
End Function

Public Sub AI_King(ByVal NPCIndex As Integer)

    '***************************************************
    'Author: ZaMa
    'Last Modification: 19/09/2010
    '***************************************************
    On Error GoTo ErrHandler

    Dim BestTarget As Integer

    Dim Action     As Byte
    
    BestTarget = KingBestTarget(NPCIndex, Action)
    
    Call KingPerformAction(NPCIndex, BestTarget, Action)
    
    Exit Sub

ErrHandler:
    LogError ("Error en AI_King. Error: " & Err.Number & " - " & Err.description)

End Sub

Private Function KingBestTarget(ByVal NPCIndex As Integer, _
                                ByRef Accion As Byte) As Integer

    '***************************************************
    'Author: ZaMa
    'Last Modification: 05/07/2010
    'Pick the best target according to the following criteria:
    '1) First try to heal status effects on allies (paralisis, poison, damaged) if any is alive
    '2) If no ally is alive, then start a melee atack to near users
    '3) If a user tries to run, then blind and make him stupid
    '4) if no one is near, then heal himself.
    '***************************************************
    On Error GoTo ErrHandler
    
    Dim BestTarget As Integer

    ' There should be more than one in order to support
    If ClanMembersAlive <> 1 Then

        ' Chooses target only if can attack
        If IntervaloPermiteAtacarNpc(NPCIndex) Then
            ' First checks if any ally is paralized in order to remove its paralizis
            BestTarget = AllyParalyzed(NPCIndex, False)

            If BestTarget <> 0 Then

                ' 20% of probability of doing it
                If RandomNumber(1, 100) < 21 Then
                    Accion = eKingAction.HealAllyParalisis

                End If

            Else
                ' Ahora comprueba si algún aliado está envenenado para eliminar su veneno.
                BestTarget = AllyPoisoned(NPCIndex)

                If BestTarget <> 0 Then
                    Accion = eKingAction.HealAllyPoison
                End If
                
                'O incinerado
                BestTarget = AllyIncinera(NPCIndex)
                
                If BestTarget <> 0 Then
                    Accion = eKingAction.HealAllyQuemaduras
                Else
                    ' If no target found, then try to heal injured allies
                    BestTarget = AllyInjured(NPCIndex)

                    If BestTarget <> 0 Then
                        Accion = eKingAction.HealAlly

                    End If

                End If

            End If

        End If
    
        ' King's alone
    Else
        
        Dim Mapa               As Integer

        Dim NPCPosX            As Integer

        Dim NPCPosY            As Integer
        
        Dim UserIndex          As Integer

        Dim Counter            As Long
        
        Dim BestTargetDistance As Integer

        Dim Distance           As Integer
        
        With Npclist(NPCIndex).Pos
            Mapa = .Map
            NPCPosX = .X
            NPCPosY = .Y

        End With
        
        Dim CounterStart As Long

        Dim CounterEnd   As Long

        Dim CounterStep  As Long
        
        Dim query()      As Collision.UUID
        
        Call ModAreas.QueryEntities(NPCIndex, ENTITY_TYPE_NPC, query, ENTITY_TYPE_PLAYER)

        ' To avoid that all attack the same target
        CounterStep = RandomNumber(0, 1)

        If CounterStep = 1 Then
            CounterStart = 1
            CounterEnd = UBound(query)
        Else
            CounterStart = UBound(query)
            CounterEnd = 1
            CounterStep = -1

        End If
        
        ' Search for the best user target
        For Counter = CounterStart To CounterEnd Step CounterStep
        
            UserIndex = query(Counter).Name
            
            'Is it in it's range of vision??
            If InVisionRange(UserIndex, NPCPosX, NPCPosY) Then

                ' Can be atacked? If it's blinded, doesn't count.
                If UserAtacable(UserIndex) And UserList(UserIndex).flags.Ceguera = 0 Then

                    ' if previous user choosen, select the better
                    If BestTarget <> 0 Then
                        ' Has priority to attack the nearest
                        Distance = UserDistance(UserIndex, NPCPosX, NPCPosY)
                        
                        If Distance < BestTargetDistance Then
                            BestTarget = UserIndex
                            BestTargetDistance = Distance

                        End If

                    Else
                        BestTarget = UserIndex
                        BestTargetDistance = UserDistance(UserIndex, NPCPosX, NPCPosY)

                    End If

                End If

            End If
                
        Next Counter
        
        ' Any target found?
        If BestTarget <> 0 Then

            ' Is reachable?
            If UserReachable(NPCIndex, BestTarget) Then
                ' Chase it
                Accion = eKingAction.ChaseTarget
            
                ' The target is trying to escape => Stop him with a massive status attack
            Else
                Accion = eKingAction.MassiveAttack

            End If
            
            ' No target found => Retreat and heal
        Else
            Accion = eKingAction.RetreatAndHeal

        End If

    End If

    KingBestTarget = BestTarget

    Exit Function

ErrHandler:
    LogError ("Error en KingBestTarget. Error: " & Err.Number & " - " & Err.description)

End Function

Private Sub KingPerformAction(ByVal NPCIndex As Integer, _
                              ByVal BestTarget As Integer, _
                              ByVal Accion As Byte)

    '***************************************************
    'Author: ZaMa
    'Last Modification: 05/07/2010
    'Performs the required action.
    '***************************************************
    On Error GoTo ErrHandler

    Dim IdleTime As Long

    With Npclist(NPCIndex)

        Select Case Accion
            
            Case eKingAction.HealAllyParalisis
                Call NpcLanzaSpellSobreNpc(NPCIndex, BestTarget, .Spells(eKingSpells.RemoveParalisis), True)
            
            Case eKingAction.HealAllyPoison
                Call NpcLanzaSpellSobreNpc(NPCIndex, BestTarget, .Spells(eKingSpells.HealPoison), True)
                
            Case eKingAction.HealAllyQuemaduras
                Call NpcLanzaSpellSobreNpc(NPCIndex, BestTarget, .Spells(eKingSpells.HealQuemaduras), True)
            
            Case eKingAction.HealAlly
                Call NpcLanzaSpellSobreNpc(NPCIndex, BestTarget, .Spells(eKingSpells.LittleHeal), True)
                
                ' If time for respawning an ally has passed, then try to do it
                If KingIdleTime > ALLY_RESPAWN_TIME Then Call KingReviveAlly(NPCIndex)
                IdleTime = KingIdleTime + 1
                
            Case eKingAction.ChaseTarget, eKingAction.MassiveAttack
                
                If Accion = eKingAction.ChaseTarget Then
                    Call GreedyWalkTo(NPCIndex, UserList(BestTarget).Pos.Map, UserList(BestTarget).Pos.X, UserList(BestTarget).Pos.Y)
                Else
                    Call NpcLanzaSpellSobreUser(NPCIndex, BestTarget, .Spells(eKingSpells.StupidityAttack), True, True)
                    Call NpcLanzaSpellSobreUser(NPCIndex, BestTarget, .Spells(eKingSpells.BlindAttack), True, True)
                    
                    ' King Message
                    Call WriteConsoleMsg(BestTarget, "El rey pretoriano te ha vuelto estupido.", FontTypeNames.FONTTYPE_FIGHT)
                    Call WriteConsoleMsg(BestTarget, "El rey pretoriano te ha vuelto ciego ", FontTypeNames.FONTTYPE_FIGHT)
                    Call WriteConsoleMsg(BestTarget, "A la distancia escuchas las siguientes palabras: Cobarde, no eres digno de luchar conmigo si escapas! ", FontTypeNames.FONTTYPE_VENENO)

                End If
            
                Call KingMeleeAttack(NPCIndex)
            
            Case eKingAction.RetreatAndHeal
                Call ReturnToCenter(NPCIndex)

                If .Stats.MinHp <> .Stats.MaxHp Then
                    Call NpcLanzaSpellSobreNpc(NPCIndex, NPCIndex, .Spells(eKingSpells.LittleHeal), True)

                End If
                
                ' If time for respawning an ally has passed, then try to do it
                If KingIdleTime > ALLY_RESPAWN_TIME Then Call KingReviveAlly(NPCIndex)
                IdleTime = KingIdleTime + 1
                
        End Select
    
    End With
    
    ' Resets if anyone is around, else increases
    KingIdleTime = IdleTime
    Exit Sub

ErrHandler:
    LogError ("Error en KingPerformAction. Error: " & Err.Number & " - " & Err.description)

End Sub

Private Sub KingMeleeAttack(NPCIndex)

    '***************************************************
    'Author: ZaMa
    'Last Modification: 19/09/2010
    'King performes a melee attack ignoring attack intervals
    '***************************************************
    On Error GoTo ErrHandler

    Dim headingloop As Byte

    Dim nPos        As WorldPos
    
    Dim UserIndex   As Integer
    
    With Npclist(NPCIndex)
    
        ' Check the four directions
        For headingloop = eHeading.SOUTH To eHeading.EAST
        
            nPos = .Pos
            
            Call HeadtoPos(headingloop, nPos)
            
            ' In map Limits?
            If InMapBounds(nPos.Map, nPos.X, nPos.Y) Then
            
                UserIndex = MapData(nPos.Map, nPos.X, nPos.Y).UserIndex

                If UserIndex > 0 Then
                    If UserAtacable(UserIndex, False) Then
                        If NpcAtacaUser(NPCIndex, UserIndex) Then
                            Call ChangeNPCChar(NPCIndex, .Char.body, .Char.Head, headingloop)

                        End If

                    End If

                End If

            End If

        Next headingloop
        
    End With
    
    Exit Sub

ErrHandler:
    LogError ("Error en KingMeleeAttack. Error: " & Err.Number & " - " & Err.description)

End Sub

Private Sub KingReviveAlly(ByVal NPCIndex As Integer)
    '***************************************************
    'Author: ZaMa
    'Last Modification: 26/09/2010
    'King tries to revive an ally.
    '***************************************************
    
    Dim AllyIndex As Integer

    Dim nPos      As WorldPos
    
    With Npclist(NPCIndex)

        ' Has to be in the center to revive
        If .Pos.X = CenterPos.X Then
            If .Pos.Y = CenterPos.Y Then
                
                ' Any ally dead?
                AllyIndex = AllyDead()

                If AllyIndex <> 0 Then
                    nPos = CenterPos
                    nPos.X = nPos.X + 1
                    Call SpawnPretorian(nPos, Pretorianos(AllyIndex).NPCAI, AllyIndex)
                    ClanMembersAlive = ClanMembersAlive + 1

                End If
                
                ' Reset iddle time
                KingIdleTime = 0
                
            End If

        End If

    End With

End Sub

Public Sub AI_Healer(ByVal NPCIndex As Integer)

    '***************************************************
    'Author: Unkown
    'Last Modification: -
    '***************************************************
    On Error GoTo ErrHandler
    
    Dim BestTarget As Integer

    Dim Accion     As Byte
    
    ' Chooses target only if can attack
    If IntervaloPermiteAtacarNpc(NPCIndex) Then
        ' First checks if any ally is paralized in order to remove its paralizis
        BestTarget = AllyParalyzed(NPCIndex, False)

        If BestTarget <> 0 Then
            Accion = eHealerAction.HealAllyParalisis
        Else
            ' Now picks the best target to attack, either user or pet
            BestTarget = HealerBestTarget(NPCIndex, Accion)
            
            ' If no target found, then try to heal injured allies
            If BestTarget = 0 Then
                BestTarget = AllyInjured(NPCIndex)
                
                ' Found an injured Ally
                If BestTarget <> 0 Then
                    Accion = eHealerAction.HealAlly

                End If
                
            End If

        End If

    End If
    
    ' Search for far paralized allies, if not better choise selected
    If BestTarget = 0 Then
        
        'if it's paralized, it has no sense doing it
        If Npclist(NPCIndex).flags.Paralizado = 0 Then
        
            BestTarget = AllyParalyzed(NPCIndex, True)

            If BestTarget <> 0 Then
                Accion = eHealerAction.RescueFarAlly

            End If
            
        End If

    End If
    
    Dim CanMove As Boolean
    
    ' Performes the best action acording to the chosen target
    Call HealerPerformAction(NPCIndex, BestTarget, Accion, CanMove)
    
    ' Moves only if in danger
    If CanMove Then Call HealerMove(NPCIndex)
    
    Exit Sub
    
ErrHandler:
    LogError ("Error en AI_Healer. Error: " & Err.Number & " - " & Err.description)

End Sub

Private Function HealerBestTarget(ByVal NPCIndex As Integer, _
                                  ByRef Accion As Byte) As Integer
    '***************************************************
    'Author: ZaMa
    'Last Modification: 09/07/2010
    'Pick the best target according to the following criteria:
    '1) "hoaxed" friends MUST be released
    '2) enemy shall be annihilated no matter what
    '3) party healing if no threats
    '***************************************************

    On Error GoTo ErrHandler

    Dim UserIndex           As Integer

    Dim Counter             As Long
    
    Dim PetIndex            As Integer
    
    Dim BestTarget          As Integer

    Dim BestTargetInvisible As Boolean
    
    Dim NpcX                As Integer

    Dim NpcY                As Integer
    
    NpcX = Npclist(NPCIndex).Pos.X
    NpcY = Npclist(NPCIndex).Pos.Y
    
    Dim CounterStart As Long

    Dim CounterEnd   As Long

    Dim CounterStep  As Long
    
    Dim query()      As Collision.UUID

    Call ModAreas.QueryEntities(NPCIndex, ENTITY_TYPE_NPC, query, ENTITY_TYPE_PLAYER)

    ' To avoid that all attack the same target
    CounterStep = RandomNumber(0, 1)

    If CounterStep = 1 Then
        CounterStart = 1
        CounterEnd = UBound(query)
    Else
        CounterStart = UBound(query)
        CounterEnd = 1
        CounterStep = -1

    End If
    
    For Counter = CounterStart To CounterEnd Step CounterStep
    
        UserIndex = query(Counter).Name
        
        ' Can be atacked (doesn't matter if invisible or hidden)?
        If UserAtacable(UserIndex, False) Then
                
            ' Checks if ther's a non-paralized user pet in range. If so, then paralize it.
            PetIndex = CheckNearUserPets(NPCIndex, UserIndex)

            If PetIndex <> 0 Then
                BestTarget = PetIndex
                Accion = eHealerAction.ParalizePet
                Exit For

            End If
                
            'Is it in it's range of vision??
            If InVisionRange(UserIndex, NpcX, NpcY) Then
                
                With UserList(UserIndex)
                    
                    ' It's Paralized?
                    If .flags.Paralizado = 0 Then
                        
                        ' It's visible?
                        If (.flags.invisible = 0) And (.flags.Oculto = 0) Then
                            
                            BestTarget = UserIndex
                            Accion = eHealerAction.ParalizeUser
                            
                        End If
                        
                        ' User Paralized, if not better target found then attack it
                    ElseIf BestTarget = 0 Then

                        BestTarget = UserIndex
                        Accion = eHealerAction.AttackUser
                        
                    End If
                    
                End With
                
            End If
        
        End If
        
    Next Counter
    
    HealerBestTarget = BestTarget
        
    Exit Function

ErrHandler:
    LogError ("Error en HealerBestTarget. Error: " & Err.Number & " - " & Err.description)

End Function

Private Sub HealerPerformAction(ByVal NPCIndex As Integer, ByVal BestTarget As Integer, ByVal Accion As Byte, _
                                ByRef CanMove As Boolean)
'***************************************************
'Author: ZaMa
'Last Modification: 09/07/2010
'Performs the required action. Determines if the npc can move after perfoming the selected action.
'***************************************************
    
    With Npclist(NPCIndex)
        Select Case Accion
            
            Case eHealerAction.HealAllyParalisis
                Call NpcLanzaSpellSobreNpc(NPCIndex, BestTarget, .Spells(eHealerSpells.RemoveParalisis), True)
                CanMove = False
                
            Case eHealerAction.ParalizePet
                Call NpcLanzaSpellSobreNpc(NPCIndex, BestTarget, .Spells(eHealerSpells.ParalizeNpc), True)
                CanMove = True
                
            Case eHealerAction.ParalizeUser
                Call SendData(SendTarget.ToNPCArea, NPCIndex, _
                    PrepareMessageChatOverHead(Hechizos(.Spells(eHealerSpells.ParalizeUser)).PalabrasMagicas, .Char.CharIndex, 0, 255, 255))
                Call NpcLanzaSpellSobreUser(NPCIndex, BestTarget, .Spells(eHealerSpells.ParalizeUser), True)
                CanMove = False
                
            Case eHealerAction.AttackUser
                Call NpcLanzaSpellSobreUser(NPCIndex, BestTarget, .Spells(eHealerSpells.Tormenta), True, True)
                CanMove = False
                
            Case eHealerAction.HealAlly
                Call NpcLanzaSpellSobreNpc(NPCIndex, BestTarget, .Spells(eHealerSpells.Heal), True)
                CanMove = True
                
            Case eHealerAction.RescueFarAlly
                Call GreedyWalkTo(NPCIndex, .Pos.Map, Npclist(BestTarget).Pos.X, Npclist(BestTarget).Pos.Y)
                CanMove = False
                 
            Case Else
                CanMove = True
                
        End Select
    End With

End Sub

Private Sub HealerMove(ByVal NPCIndex As Integer)
    '***************************************************
    'Author: ZaMa
    'Last Modification: 09/07/2010
    'Doesn't move after casting some spells. Try to avoid swordmasters.
    '***************************************************
    
    With Npclist(NPCIndex)
        
        ' Can't move if paralized
        If .flags.Paralizado = 1 Then Exit Sub
    
        ' Only applies to clan map
        If .Pos.Map <> CenterPos.Map Then Exit Sub
        
        ' if nobody's near or it's idle then returns near the king
        Call ReturnToCenter(NPCIndex)
        
        ' Runs away from close users
        If TargetClose(.Pos) Then
            Call ReturnToCenter(NPCIndex)

        End If
        
    End With

End Sub

Public Sub AI_SpellCaster(ByVal NPCIndex As Integer)

    '***************************************************
    'Author: ZaMa
    'Last Modification: -
    '***************************************************
    On Error GoTo ErrHandler
    
    Dim BestTarget As Integer

    Dim Accion     As Byte
    
    ' It sacrifices itself giving a fatal blow if close to death
    If SpellCasterSacrifice(NPCIndex) Then Exit Sub
    
    ' If healthy then choose best target
    BestTarget = SpellCasterBestTarget(NPCIndex, Accion)
    
    ' Performes the best action acoording to the chosen target
    If BestTarget <> 0 Then
        Call SpellCasterPerformAction(NPCIndex, BestTarget, Accion)

    End If
    
    ' Moves only if in danger
    Call SpellCasterMove(NPCIndex, BestTarget)
    
    Exit Sub
    
ErrHandler:
    LogError ("Error en AI_SpellCaster. Error: " & Err.Number & " - " & Err.description)

End Sub

Private Function SpellCasterSacrifice(ByVal NPCIndex As Integer) As Boolean
    '***************************************************
    'Author: ZaMa
    'Last Modification: 05/07/2010
    'Checks if the required conditions to sacrify are fullfiled, if so then does it.
    'The wand break counter is stored in the boat slot, asuming it'd never be used for any other purpose.
    '***************************************************
    
    With Npclist(NPCIndex)
    
        If .Stats.MinHp <= START_COUNTDOWN_MIN_HP Then
            Call SendData(SendTarget.ToNPCArea, NPCIndex, PrepareMessageCreateFX(.Char.CharIndex, FXIDs.FXMEDITARGRANDE, INFINITE_LOOPS))
        
            If IntervaloPermiteAtacarNpc(NPCIndex) Then
                
                WandBreakCounter = WandBreakCounter - 1
                Call SpellCasterDestroyWand(NPCIndex, eSpellCasterSpells.Apocalipsis)
                
                '  Wand broken => dies
                If WandBreakCounter = 0 Then Call MuereNpc(NPCIndex, 0)
                
            End If
            
            SpellCasterSacrifice = True
            
        Else

            'Restore wand break counter
            If WandBreakCounter <> MAX_WAND_BREAK_VALUE Then
                WandBreakCounter = MAX_WAND_BREAK_VALUE
                Call SendData(SendTarget.ToNPCArea, NPCIndex, PrepareMessageCreateFX(.Char.CharIndex, 0, 0))

            End If

        End If

    End With
    
End Function

Private Function SpellCasterBestTarget(ByVal NPCIndex As Integer, _
                                       ByRef Accion As Byte) As Integer
    '***************************************************
    'Author: ZaMa
    'Last Modification: 05/07/2010
    'Pick the best target according to the following criteria:
    '1) Invisible enemies can be detected sometimes
    '2) A wizard's mission is background spellcasting attack
    '***************************************************

    On Error GoTo ErrHandler

    Dim UserIndex           As Integer

    Dim Counter             As Long
    
    Dim BestTarget          As Integer

    Dim NpcX                As Integer

    Dim NpcY                As Integer

    Dim BestTargetInvisible As Boolean
        
    NpcX = Npclist(NPCIndex).Pos.X
    NpcY = Npclist(NPCIndex).Pos.Y
    
    Dim CounterStart As Long

    Dim CounterEnd   As Long

    Dim CounterStep  As Long
    
    Dim query()      As Collision.UUID

    Call ModAreas.QueryEntities(NPCIndex, ENTITY_TYPE_NPC, query, ENTITY_TYPE_PLAYER)
    
    For Counter = 0 To UBound(query)
    
        UserIndex = query(Counter).Name
        
        'Is it in it's range of vision??
        If InVisionRange(UserIndex, NpcX, NpcY) Then
            
            ' Can be atacked?
            If UserAtacable(UserIndex, False) Then
                
                With UserList(UserIndex)

                    ' It's invisible?
                    If (.flags.invisible = 1) Or (.flags.Oculto = 1) Then
                        
                        ' Try to remove invisibility
                        If (RandomNumber(1, 100) <= 35) Then
                            BestTarget = UserIndex
                            Accion = eSpellCasterAction.RemoveInvi
                            Exit For

                        End If
                        
                        ' Paralized invisible users are good target!
                        If .flags.Paralizado = 1 Then
                            BestTarget = UserIndex
                            BestTargetInvisible = True
                            Accion = eSpellCasterAction.Attack

                        End If
                    
                        ' Visible but paralized?
                    ElseIf (.flags.Paralizado = 1) Then
                        
                        ' If not found an invisible and paralized target, then it's a good one
                        If Not BestTargetInvisible Or BestTarget = 0 Then
                            BestTarget = UserIndex
                            Accion = eSpellCasterAction.Attack

                        End If
                    
                        ' Visible and can move, if not better found then choose it
                    ElseIf BestTarget = 0 Then

                        BestTarget = UserIndex
                        Accion = eSpellCasterAction.Attack
                        
                    End If
                    
                End With
                
            End If
        
        End If
        
    Next Counter
    
    SpellCasterBestTarget = BestTarget
    
    Exit Function

ErrHandler:
    LogError ("Error en SpellCasterBestTarget. Error: " & Err.Number & " - " & Err.description)

End Function

Private Sub SpellCasterPerformAction(ByVal NPCIndex As Integer, _
                                     ByVal BestTarget As Integer, _
                                     ByVal Accion As Byte)
    '***************************************************
    'Author: ZaMa
    'Last Modification: 05/07/2010
    'Performs the required action.
    '***************************************************
    
    With Npclist(NPCIndex)

        Select Case Accion
            
            Case eSpellCasterAction.Attack
                Call NpcLanzaSpellSobreUser(NPCIndex, BestTarget, .Spells(eSpellCasterSpells.Apocalipsis), True, True)
            
            Case eSpellCasterAction.RemoveInvi
                Call NpcLanzaSpellSobreUser(NPCIndex, BestTarget, .Spells(eSpellCasterSpells.RemoInvi), True, True)
                
        End Select

    End With
    
End Sub

Private Sub SpellCasterMove(ByVal NPCIndex As Integer, ByVal BestTarget As Integer)
    '***************************************************
    'Author: ZaMa
    'Last Modification: 05/07/2010
    'Doesn't move unless nobody is near or there is no best target.
    '***************************************************
    
    With Npclist(NPCIndex)
        
        ' Can't move if paralized
        If .flags.Paralizado = 1 Then Exit Sub
        
        ' Only applies to clan map
        If .Pos.Map <> CenterPos.Map Then Exit Sub
        
        ' if nobody's near or it's idle then returns near the king
        If (BestTarget = 0) And (IntervaloPermiteAtacarNpc(NPCIndex)) Then Call ReturnToCenter(NPCIndex)
        
        ' Runs away from close users
        If TargetClose(.Pos) Then
            Call ReturnToCenter(NPCIndex)

        End If
        
    End With

End Sub

Private Sub SpellCasterDestroyWand(ByVal NPCIndex As Integer, ByVal spellIndex As Integer)

    '***************************************************
    'Author: Unkown
    'Last Modification: -
    '***************************************************
    On Error GoTo ErrHandler

    ''sonidos: 30 y 32, y no los cambien sino termina siendo muy chistoso...
    ''Para el FX utiliza el del hechizos(SpellIndex)
    
    With Npclist(NPCIndex)

        Select Case WandBreakCounter

            Case 5
                Call SendData(SendTarget.ToNPCArea, NPCIndex, PrepareMessageChatOverHead("Rahma", .Char.CharIndex, 0, 255, 0))
                Call SendData(SendTarget.ToNPCArea, NPCIndex, PrepareMessagePlayWave(SONIDO_DRAGON_VIVO, .Pos.X, .Pos.Y))

            Case 4
                Call SendData(SendTarget.ToNPCArea, NPCIndex, PrepareMessageChatOverHead("vortax", .Char.CharIndex, 0, 255, 0))
                Call SendData(SendTarget.ToNPCArea, NPCIndex, PrepareMessagePlayWave(SONIDO_DRAGON_VIVO, .Pos.X, .Pos.Y))

            Case 3
                Call SendData(SendTarget.ToNPCArea, NPCIndex, PrepareMessageChatOverHead("Zill", .Char.CharIndex, 0, 255, 0))
                Call SendData(SendTarget.ToNPCArea, NPCIndex, PrepareMessagePlayWave(SONIDO_DRAGON_VIVO, .Pos.X, .Pos.Y))

            Case 2
                Call SendData(SendTarget.ToNPCArea, NPCIndex, PrepareMessageChatOverHead("yaka E'nta", .Char.CharIndex, 0, 255, 0))
                Call SendData(SendTarget.ToNPCArea, NPCIndex, PrepareMessagePlayWave(SONIDO_DRAGON_VIVO, .Pos.X, .Pos.Y))

            Case 1
                Call SendData(SendTarget.ToNPCArea, NPCIndex, PrepareMessageChatOverHead("Korata!!", .Char.CharIndex, 0, 255, 0))
                Call SendData(SendTarget.ToNPCArea, NPCIndex, PrepareMessagePlayWave(SONIDO_DRAGON_VIVO, .Pos.X, .Pos.Y))

            Case 0
                Call SendData(SendTarget.ToNPCArea, NPCIndex, PrepareMessageChatOverHead(vbNullString, .Char.CharIndex, 0, 255, 0))
                Call SendData(SendTarget.ToNPCArea, NPCIndex, PrepareMessagePlayWave(SONIDO_DRAGON_VIVO, .Pos.X, .Pos.Y))
                
                Call DealWandDamage(NPCIndex, .Spells(spellIndex))

        End Select
        
    End With
    
    Exit Sub

ErrHandler:
    LogError ("Error en SpellCasterDestroyWand. Error: " & Err.Number & " - " & Err.description)

End Sub

Private Sub DealWandDamage(ByVal NPCIndex As Integer, ByVal spellIndex As Integer)

    '***************************************************
    'Author: ZaMa
    'Last Modification: 17/06/2011
    'Description: Deals massive wand damage.
    '17/06/2011: Amraphen - User's hp is updated once he gets hit by wand damage.
    '***************************************************
    On Error GoTo ErrHandler

    Dim UserIndex As Integer

    Dim Counter   As Long
    
    Dim Distancia As Integer

    Dim Danio     As Integer
    
    Dim X         As Integer

    Dim Y         As Integer
    
    With Npclist(NPCIndex)
        
        X = .Pos.X
        Y = .Pos.Y
        
        Dim query() As Collision.UUID

        Call ModAreas.QueryEntities(NPCIndex, ENTITY_TYPE_NPC, query, ENTITY_TYPE_PLAYER)

        For Counter = 1 To UBound(query)
            
            If UserIndex <> 0 Then
            
                With UserList(UserIndex)
                    
                    If UserAtacable(UserIndex, False, False) Then
                        
                        Distancia = UserDistance(UserIndex, X, Y)
                        Danio = Abs(Int(880 / (Distancia ^ (3 / 7))))
                        
                        .Stats.MinHp = .Stats.MinHp - Danio
                        
                        Call WriteConsoleMsg(UserIndex, Npclist(NPCIndex).Name & " te ha quitado " & Danio & " puntos de vida al romper su vara.", FontTypeNames.FONTTYPE_FIGHT)
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Hechizos(spellIndex).WAV, .Pos.X, .Pos.Y))
                        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, Hechizos(spellIndex).FXgrh, Hechizos(spellIndex).loops))
                        
                        If .Stats.MinHp < 1 Then
                            .Stats.MinHp = 0
                            Call UserDie(UserIndex)
                        Else
                            'Updates user's hp:
                            Call WriteUpdateHP(UserIndex)

                            ' Also hit pets
                            If .NroMascotas > 0 Then Call DealWandDamageToPets(NPCIndex, UserIndex, spellIndex)
                            
                        End If
                        
                    End If
                    
                End With
            
            End If
            
        Next Counter
        
    End With
    
    Exit Sub
    
ErrHandler:
    LogError ("Error en DealWandDamage. Error: " & Err.Number & " - " & Err.description)

End Sub

Private Sub DealWandDamageToPets(ByVal NPCIndex As Integer, _
                                 ByVal UserIndex As Integer, _
                                 ByVal spellIndex As Integer)

    '***************************************************
    'Author: ZaMa
    'Last Modification: 26/09/2010
    'Deals massive wand damage to user pets.
    '***************************************************
    On Error GoTo ErrHandler

    Dim PetIndex   As Integer

    Dim PetCounter As Long

    Dim Distancia  As Integer

    Dim Danio      As Integer
    
    Dim NpcX       As Integer

    Dim NpcY       As Integer
    
    NpcX = Npclist(NPCIndex).Pos.X
    NpcY = Npclist(NPCIndex).Pos.Y
    
    For PetCounter = 1 To MAXMASCOTAS
                                    
        PetIndex = UserList(UserIndex).MascotasIndex(PetCounter)

        If PetIndex > 0 Then

            With Npclist(PetIndex)
           
                Distancia = NpcDistance(PetIndex, NpcX, NpcY)
                Danio = Abs(Int(880 / (Distancia ^ (3 / 7))))

                .Stats.MinHp = .Stats.MinHp - Danio
                
                Call SendData(SendTarget.ToNPCArea, PetIndex, PrepareMessagePlayWave(Hechizos(spellIndex).WAV, .Pos.X, .Pos.Y))
                Call SendData(SendTarget.ToNPCArea, PetIndex, PrepareMessageCreateFX(.Char.CharIndex, Hechizos(spellIndex).FXgrh, Hechizos(spellIndex).loops))
                
                If .Stats.MinHp < 1 Then
                    .Stats.MinHp = 0
                    Call MuereNpc(PetIndex, 0)

                End If
                
            End With

        End If
        
    Next PetCounter
    
    Exit Sub
    
ErrHandler:
    LogError ("Error en DealWandDamageToPets. Error: " & Err.Number & " - " & Err.description)

End Sub

Public Sub AI_SwordMaster(ByVal NPCIndex As Integer)

    '***************************************************
    'Author: ZaMa
    'Last Modification: -
    '***************************************************
    On Error GoTo ErrHandler

    Dim BestTarget As Integer
    
    ' First choose best target
    BestTarget = SwordMasterBestTarget(NPCIndex)
    
    ' Then moves, close to target user if in range
    Call SwordMasterMove(NPCIndex, BestTarget)
    
    ' Finally, attacks sorrounding targets
    Call SwordMasterAttack(NPCIndex)
    
    Exit Sub

ErrHandler:
    LogError ("Error en AI_SwordMaster. Error: " & Err.Number & " - " & Err.description)

End Sub

Private Function SwordMasterBestTarget(ByVal NPCIndex As Integer) As Integer
    '***************************************************
    'Author: ZaMa
    'Last Modification: 26/06/2010
    'Picks the best target according to the following criteria:
    '1) The nearest enemy will be attacked.
    '***************************************************

    On Error GoTo ErrHandler

    Dim UserIndex          As Integer

    Dim Counter            As Long
    
    Dim BestTarget         As Integer

    Dim BestTargetDistance As Integer
    
    Dim NpcX               As Integer

    Dim NpcY               As Integer

    Dim Distance           As Integer
    
    NpcX = Npclist(NPCIndex).Pos.X
    NpcY = Npclist(NPCIndex).Pos.Y
    
    Dim CounterStart As Long

    Dim CounterEnd   As Long

    Dim CounterStep  As Long
    
    Dim query()      As Collision.UUID

    Call ModAreas.QueryEntities(NPCIndex, ENTITY_TYPE_NPC, query, ENTITY_TYPE_PLAYER)

    ' To avoid that all attack the same target
    CounterStep = RandomNumber(0, 1)

    If CounterStep = 1 Then
        CounterStart = 1
        CounterEnd = UBound(query)
    Else
        CounterStart = UBound(query)
        CounterEnd = 1
        CounterStep = -1

    End If
    
    For Counter = CounterStart To CounterEnd Step CounterStep
    
        UserIndex = query(Counter).Name
        
        'Is it in it's range of vision??
        If InVisionRange(UserIndex, NpcX, NpcY) Then
            
            ' Can be atacked?
            If UserAtacable(UserIndex) Then

                ' Checks limits in order to avoid to separate it from the rest of the clan
                If UserReachable(NPCIndex, UserIndex) Then
                
                    If BestTarget <> 0 Then
                        
                        ' Has priority to attack the nearest
                        Distance = UserDistance(UserIndex, NpcX, NpcY)
                        
                        If Distance < BestTargetDistance Then
                            BestTarget = UserIndex
                            BestTargetDistance = Distance

                        End If
                        
                    Else
                        BestTarget = UserIndex
                        BestTargetDistance = UserDistance(UserIndex, NpcX, NpcY)

                    End If
                
                End If
            
            End If
        
        End If
        
    Next Counter
    
    SwordMasterBestTarget = BestTarget
    
    Exit Function

ErrHandler:
    LogError ("Error en SwordMasterBestTarget. Error: " & Err.Number & " - " & Err.description)

End Function

Private Sub SwordMasterMove(ByVal NPCIndex As Integer, ByVal BestTarget As Integer)
    '***************************************************
    'Author: ZaMa
    'Last Modification: 24/06/2010
    'La idea es que no lo "alejen" del rey y despues queden
    'lejos de la batalla cuando matan a un enemigo o este
    'sale del area de combate (tipica forma de separar un clan)
    '***************************************************

    On Error GoTo ErrHandler

    With Npclist(NPCIndex)
    
        ' If paralized can't move
        If .flags.Paralizado = 1 Then Exit Sub
        
        ' Only applies to clan map
        If .Pos.Map <> CenterPos.Map Then Exit Sub
        
        ' If it's far from protection then returns to center
        If FarFromTeam(NPCIndex) Then
            Call ReturnToCenter(NPCIndex)
            Exit Sub

        End If
        
        ' If someone's in range, then go close to him
        If BestTarget > 0 Then
            
            With UserList(BestTarget)
                Call GreedyWalkTo(NPCIndex, CenterPos.Map, .Pos.X, .Pos.Y)

            End With
            
            'Returns to center if no target found
        Else
            Call ReturnToCenter(NPCIndex)

        End If
        
    End With
    
    Exit Sub
    
ErrHandler:
    LogError ("Error en SwordMasterMove. Error: " & Err.Number & " - " & Err.description)

End Sub

Private Sub SwordMasterAttack(ByVal NPCIndex As Integer)
    '***************************************************
    'Author: ZaMa
    'Last Modification: 26/06/2010
    'Searchs and Attacks the sorrounding target
    '***************************************************

    On Error GoTo ErrHandler

    Dim headingloop As Byte

    Dim NpcPos      As WorldPos

    Dim UserIndex   As Integer
    
    With Npclist(NPCIndex)
    
        ' Check the four directions
        For headingloop = eHeading.SOUTH To eHeading.EAST
        
            NpcPos = .Pos
            Call HeadtoPos(headingloop, NpcPos)
            
            UserIndex = MapData(NpcPos.Map, NpcPos.X, NpcPos.Y).UserIndex
            
            If UserIndex > 0 Then
                If UserAtacable(UserIndex, False) Then
                    If NpcAtacaUser(NPCIndex, UserIndex) Then
                        Call ChangeNPCChar(NPCIndex, .Char.body, .Char.Head, headingloop)
                        Exit For

                    End If

                End If

            End If
            
        Next headingloop
        
    End With

    Exit Sub
    
ErrHandler:
    LogError ("Error en SwordMasterAttack. Error: " & Err.Number & " - " & Err.description)

End Sub

Public Sub AI_Shooter(ByVal NPCIndex As Integer)

    '***************************************************
    'Author: ZaMa
    'Last Modification: -
    '***************************************************
    Dim BestTarget As Integer
    
    ' First choose best target
    BestTarget = ShooterBestTarget(NPCIndex)
    
    ' Found a target => attack!
    If BestTarget > 0 Then
        ' Attack with arrow
        Call NpcLanzaSpellSobreUser(NPCIndex, BestTarget, Npclist(NPCIndex).Spells(1))

    End If
    
    ' Then moves, close to target user if attacking one
    Call ShooterMove(NPCIndex, BestTarget)
    
    Exit Sub
    
ErrHandler:
    LogError ("Error en AI_Shooter. Error: " & Err.Number & " - " & Err.description)

End Sub

Private Function ShooterBestTarget(ByVal NPCIndex As Integer) As Integer
    '***************************************************
    'Author: ZaMa
    'Last Modification: 24/06/2010
    'Picks the best target according to the following criteria:
    '1) Magic clases ARE dangerous, but they are weak too, they're
    '   our primary target
    '2) in any other case, our nearest enemy will be attacked
    '***************************************************

    On Error GoTo ErrHandler

    Dim UserIndex          As Integer

    Dim Counter            As Long
    
    Dim BestTarget         As Integer

    Dim BestTargetDistance As Integer
    
    Dim NpcX               As Integer

    Dim NpcY               As Integer

    Dim Distance           As Integer
        
    NpcX = Npclist(NPCIndex).Pos.X
    NpcY = Npclist(NPCIndex).Pos.Y
        
    Dim CounterStart As Long

    Dim CounterEnd   As Long

    Dim CounterStep  As Long
    
    Dim query()      As Collision.UUID

    Call ModAreas.QueryEntities(NPCIndex, ENTITY_TYPE_NPC, query, ENTITY_TYPE_PLAYER)

    ' To avoid that all attack the same target
    CounterStep = RandomNumber(0, 1)

    If CounterStep = 1 Then
        CounterStart = 1
        CounterEnd = UBound(query)
    Else
        CounterStart = UBound(query)
        CounterEnd = 1
        CounterStep = -1

    End If
    
    For Counter = CounterStart To CounterEnd Step CounterStep
    
        UserIndex = query(Counter).Name
        
        'Is it in it's range of vision??
        If InVisionRange(UserIndex, NpcX, NpcY) Then
            
            ' Can be atacked?
            If UserAtacable(UserIndex) Then
                
                ' Has Priority to attack magic clases
                If EsClaseMagica(UserIndex) Then
                    BestTarget = UserIndex
                    Exit For
                Else

                    If BestTarget <> 0 Then
                        
                        ' Has priority to attack the nearest
                        Distance = UserDistance(UserIndex, NpcX, NpcY)
                        
                        If Distance < BestTargetDistance Then
                            BestTarget = UserIndex
                            BestTargetDistance = Distance

                        End If
                        
                    Else
                        BestTarget = UserIndex
                        BestTargetDistance = UserDistance(UserIndex, NpcX, NpcY)

                    End If
                    
                End If
            
            End If
        
        End If
        
    Next Counter
    
    ShooterBestTarget = BestTarget
        
    Exit Function

ErrHandler:
    LogError ("Error en ShooterBestTarget. Error: " & Err.Number & " - " & Err.description)

End Function

Private Sub ShooterMove(ByVal NPCIndex As Integer, ByVal BestTarget As Integer)
    '***************************************************
    'Author: ZaMa
    'Last Modification: 24/06/2010
    'Vamos a setear el hold on del cazador en el medio entre el rey
    'y el atacante. De esta manera se lo podra atacar aun asi esta lejos
    'pero sin alejarse del rango de los an hoax vorps de los
    'clerigos o rey. A menos q este paralizado, claro
    '***************************************************

    On Error GoTo ErrHandler

    With Npclist(NPCIndex)
    
        ' If paralized can't move
        If .flags.Paralizado = 1 Then Exit Sub
        
        ' Only applies to clan map
        If .Pos.Map <> CenterPos.Map Then Exit Sub

        ' If it's far from protection then returns to center
        If FarFromTeam(NPCIndex) Then
            Call ReturnToCenter(NPCIndex)
            Exit Sub

        End If
        
        ' Don't go too far from center
        If Abs(CenterPos.X - .Pos.X) > 4 Or Abs(CenterPos.Y - .Pos.Y) > 4 Then
            Call ReturnToCenter(NPCIndex)
            Exit Sub

        End If
        
        ' If has attacked someone, then go close to him
        If BestTarget > 0 Then
            
            Call GreedyWalkTo(NPCIndex, CenterPos.Map, CenterPos.X + ((UserList(BestTarget).Pos.X - CenterPos.X) / 2), CenterPos.Y + ((UserList(BestTarget).Pos.Y - CenterPos.Y) / 2))
            Exit Sub
        Else
        
            ' Search for aproaching enemies
            Dim Counter   As Long

            Dim UserIndex As Integer
            
            Dim query()   As Collision.UUID

            Call ModAreas.QueryEntities(NPCIndex, ENTITY_TYPE_NPC, query, ENTITY_TYPE_PLAYER)

            For Counter = 1 To UBound(query)
            
                UserIndex = query(Counter).Name
                
                'Is it in extended range of vision from center position??
                If InVisionRange(UserIndex, CenterPos.X, CenterPos.Y, True) Then
                    
                    ' Can be atacked?
                    If UserAtacable(UserIndex) Then
                        
                        With UserList(UserIndex)
                            Call GreedyWalkTo(NPCIndex, CenterPos.Map, .Pos.X, .Pos.Y)

                        End With
                        
                        Exit Sub

                    End If
                    
                End If

            Next Counter

        End If
        
        'Returns to center if no target found
        If IntervaloPermiteAtacarNpc(NPCIndex) Then Call ReturnToCenter(NPCIndex)
         
    End With
    
    Exit Sub
    
ErrHandler:
    LogError ("Error en ShooterMove. Error: " & Err.Number & " - " & Err.description)

End Sub

Private Sub AI_Thief(ByVal NPCIndex As Integer)

    '***************************************************
    'Author: ZaMa
    'Last Modification: -
    '***************************************************
    On Error GoTo ErrHandler

    Dim BestTarget As Integer

    Dim Action     As Byte
    
    ' First choose best target
    BestTarget = ThiefBestTarget(NPCIndex, Action)
    
    ' Perform's thief action acording to its priorities
    Call ThiefPerfomAction(NPCIndex, BestTarget, Action)
    
    Exit Sub

ErrHandler:
    LogError ("Error en AI_Thief. Error: " & Err.Number & " - " & Err.description)

End Sub

Private Function ThiefBestTarget(ByVal NPCIndex As Integer, _
                                 ByRef Action As Byte) As Integer
    '***************************************************
    'Author: ZaMa
    'Last Modification: 26/06/2010
    'Picks the best target according to the following criteria:
    '1) Has priority to attack hunters.
    '2) Targets the nearer user.
    '***************************************************

    On Error GoTo ErrHandler

    Dim UserIndex          As Integer

    Dim Counter            As Long
    
    Dim BestTarget         As Integer

    Dim BestTargetDistance As Integer
    
    Dim NpcX               As Integer

    Dim NpcY               As Integer

    Dim Distance           As Integer
    
    NpcX = Npclist(NPCIndex).Pos.X
    NpcY = Npclist(NPCIndex).Pos.Y
    
    Dim CounterStart As Long

    Dim CounterEnd   As Long

    Dim CounterStep  As Long
    
    Dim query()      As Collision.UUID

    Call ModAreas.QueryEntities(NPCIndex, ENTITY_TYPE_NPC, query, ENTITY_TYPE_PLAYER)

    ' To avoid that all attack the same target
    CounterStep = RandomNumber(0, 1)

    If CounterStep = 1 Then
        CounterStart = 1
        CounterEnd = UBound(query)
    Else
        CounterStart = UBound(query)
        CounterEnd = 1
        CounterStep = -1

    End If
    
    For Counter = CounterStart To CounterEnd Step CounterStep
    
        UserIndex = query(Counter).Name
        
        'Is it in it's range of vision??
        If InVisionRange(UserIndex, NpcX, NpcY) Then
            
            ' Can be atacked?
            If UserAtacable(UserIndex, False) Then

                ' Checks limits in order to avoid to separate it from the rest of the clan
                If UserReachable(NPCIndex, UserIndex) Then
                
                    ' Is it a hunter? => Has priority
                    If UserList(UserIndex).clase = eClass.Hunter Then
                        BestTarget = UserIndex
                        BestTargetDistance = UserDistance(UserIndex, NpcX, NpcY)
                        Exit For

                    End If
                    
                    If BestTarget <> 0 Then
                        
                        ' Has priority to attack the nearest
                        Distance = UserDistance(UserIndex, NpcX, NpcY)
                        
                        If Distance < BestTargetDistance Then
                            BestTarget = UserIndex
                            BestTargetDistance = Distance

                        End If
                        
                    Else
                        BestTarget = UserIndex
                        BestTargetDistance = UserDistance(UserIndex, NpcX, NpcY)

                    End If
                
                End If
            
            End If
        
        End If
        
    Next Counter
    
    ' Couldn't find anyone
    If BestTarget = 0 Then
        Action = eThiefAction.None
    Else

        ' if close to target => Snatch/Steal
        If BestTargetDistance <= THIEF_STEAL_DISTANCE Then
            Action = eThiefAction.Steal
        
            ' Too far => Attack while aproaching
        Else
            Action = eThiefAction.Attack

        End If

    End If
    
    ThiefBestTarget = BestTarget
    
    Exit Function

ErrHandler:
    LogError ("Error en ThiefBestTarget. Error: " & Err.Number & " - " & Err.description)

End Function

Private Sub ThiefPerfomAction(ByVal NPCIndex As Integer, _
                              ByVal BestTarget As Integer, _
                              ByVal Action As Byte)

    '***************************************************
    'Author: ZaMa
    'Last Modification: 28/10/2010
    '28/10/2010: ZaMa - Now thief doesn't steal/snatch admins.
    '***************************************************
    On Error GoTo ErrHandler
    
    Dim WeaponEqpSlot   As Byte

    Dim MunicionEqpSlot As Byte
    
    With Npclist(NPCIndex)
        
        Select Case Action
            
            Case eThiefAction.Attack, eThiefAction.Steal
                
                ' Can't do anything
                If IntervaloPermiteAtacarNpc(NPCIndex) Then
                
                    ' Attack with arrow
                    If Action = eThiefAction.Attack Then
                        Call NpcLanzaSpellSobreUser(NPCIndex, BestTarget, .Spells(eThiefSpells.Arrow), False, True)
                        
                        ' Can paralize
                        If RandomNumber(1, 100) < 10 Then
                            Call NpcLanzaSpellSobreUser(NPCIndex, BestTarget, .Spells(eThiefSpells.Paralisis), False, True)
                            Call WriteConsoleMsg(BestTarget, "" & .Name & " te ha paralizado con su golpe!!", FontTypeNames.FONTTYPE_VENENO)
                        
                            ' Can unequip weapon
                        ElseIf RandomNumber(1, 100) < 22 Then
                            
                            WeaponEqpSlot = UserList(BestTarget).Invent.WeaponEqpSlot

                            If WeaponEqpSlot <> 0 Then Call Desequipar(BestTarget, WeaponEqpSlot)
                            
                        End If
                        
                        ' Try to snatch/Steal (except for admins)
                    ElseIf Not EsGm(BestTarget) Then
                        
                        ' Can snatch weapon or arrows
                        If RandomNumber(1, 100) < 16 Then
                            
                            WeaponEqpSlot = UserList(BestTarget).Invent.WeaponEqpSlot

                            If WeaponEqpSlot <> 0 Then
                                
                                If SnatchItem(NPCIndex, BestTarget, WeaponEqpSlot) Then
                                    Call WriteConsoleMsg(BestTarget, "" & .Name & " te ha arrebatado tu arma!!", FontTypeNames.FONTTYPE_VENENO)

                                End If
                                
                            Else
                                MunicionEqpSlot = UserList(BestTarget).Invent.MunicionEqpSlot

                                If MunicionEqpSlot <> 0 Then
                                
                                    If SnatchItem(NPCIndex, BestTarget, MunicionEqpSlot) Then
                                        Call WriteConsoleMsg(BestTarget, "" & .Name & " te ha arrebatado tus municiones!!", FontTypeNames.FONTTYPE_VENENO)

                                    End If
                                    
                                End If

                            End If
                        
                            ' Can steal the items of a slot
                        ElseIf RandomNumber(1, 100) < 16 Then
                        
                            Dim Slot As Byte

                            Slot = RandomNumber(1, UserList(BestTarget).CurrentInventorySlots)
                            
                            If ThiefStealITem(NPCIndex, BestTarget, Slot) Then
                                Call WriteConsoleMsg(BestTarget, "" & .Name & " te esta robando!!", FontTypeNames.FONTTYPE_VENENO)

                            End If
                            
                        End If

                    End If
                    
                End If
                
                ' Has chances of becoming invisible (if visible)
                If .flags.invisible = 0 Then
                    
                    If RandomNumber(1, 100) < 25 Then Call ThiefTurnInvisible(NPCIndex, True)
                
                    ' Is invisible, can turn visible.
                Else

                    If RandomNumber(1, 100) < 13 Then Call ThiefTurnInvisible(NPCIndex, False)

                End If
                
                ' Get closer to target
                Call GreedyWalkTo(NPCIndex, CenterPos.Map, UserList(BestTarget).Pos.X, UserList(BestTarget).Pos.Y)
                
            Case eThiefAction.None
                
                ' Turns visible if no threat
                If .flags.invisible = 1 Then Call ThiefTurnInvisible(NPCIndex, False)
                
                Call ReturnToCenter(NPCIndex)
                
        End Select
        
    End With

    Exit Sub

ErrHandler:
    LogError ("Error en ThiefPerfomAction. Error: " & Err.Number & " - " & Err.description)

End Sub

Private Function SnatchItem(ByVal NPCIndex As Integer, _
                            ByVal targetIndex As Integer, _
                            ByVal Slot As Byte) As Boolean

    '***************************************************
    'Author: ZaMa
    'Last Modification: 25/09/2010
    'Snatchs slot items from user. No validations are made bacause it's either weapon or munition.
    '***************************************************
    On Error GoTo ErrHandler

    Dim DropObj As obj

    Dim nPos    As WorldPos
    
    With UserList(targetIndex)
        DropObj.Amount = .Invent.Object(Slot).Amount
        DropObj.ObjIndex = .Invent.Object(Slot).ObjIndex
    
        ' Search for suitable place to drop item
        Call Tilelibre(.Pos, nPos, DropObj, False, True)
        
        ' Found any?
        If nPos.X <> 0 And nPos.Y <> 0 Then
            ' Drop it
            Call MakeObj(DropObj, nPos.Map, nPos.X, nPos.Y)
            Call QuitarUserInvItem(targetIndex, Slot, DropObj.Amount)
            Call UpdateUserInv(False, targetIndex, Slot)
            
            ' Suceed
            SnatchItem = True

        End If

    End With
    
    Exit Function
    
ErrHandler:
    LogError ("Error en SnatchItem. Error: " & Err.Number & " - " & Err.description)

End Function

Private Function ThiefStealITem(ByVal NPCIndex As Integer, _
                                ByVal targetIndex As Integer, _
                                ByVal Slot As Byte) As Boolean

    '***************************************************
    'Author: ZaMa
    'Last Modification: 25/09/2010
    'Steal slot items from user. If no space in thief's inventory, drop it on floor.
    '22/10/2010: ZaMa - Now thief doesn't try to stel slots with no items in it.
    '***************************************************
    Dim StolenObj As obj

    Dim NroItems  As Integer
    
    ' If doesn't have any item, the doesn't do anything
    If UserList(targetIndex).Invent.Object(Slot).ObjIndex = 0 Then Exit Function
    
    ' If not stealable, then don't do anything
    If Not ObjEsRobable(targetIndex, Slot) Then Exit Function

    ' If thief's inventory is full, then snatch it
    NroItems = Npclist(NPCIndex).Invent.NroItems

    If NroItems = 20 Then
        
        ThiefStealITem = SnatchItem(NPCIndex, targetIndex, Slot)
             
        ' Steal it
    Else
        
        StolenObj.Amount = UserList(targetIndex).Invent.Object(Slot).Amount
        StolenObj.ObjIndex = UserList(targetIndex).Invent.Object(Slot).ObjIndex
        
        ' Add it to Thief inventory
        With Npclist(NPCIndex)
            NroItems = NroItems + 1
            .Invent.Object(NroItems).Amount = StolenObj.Amount
            .Invent.Object(NroItems).ObjIndex = StolenObj.ObjIndex
            .Invent.NroItems = NroItems

        End With
        
        ' Take it from user inventory
        Call QuitarUserInvItem(targetIndex, Slot, StolenObj.Amount)
        Call UpdateUserInv(False, targetIndex, Slot)
        
        ThiefStealITem = True
        
    End If
    
End Function

Private Sub ThiefTurnInvisible(ByVal NPCIndex As Integer, ByVal TurnVisible As Boolean)
    '***************************************************
    'Author: ZaMa
    'Last Modification: 26/09/2010
    'Turn thief visible/Invisible.
    '***************************************************
    
    With Npclist(NPCIndex)
    
        If TurnVisible Then
            .flags.invisible = 1
        Else
            .flags.invisible = 0

        End If
        
        Call SendData(SendTarget.ToNPCArea, NPCIndex, PrepareMessageSetInvisible(.Char.CharIndex, TurnVisible))
    
    End With
    
End Sub

Private Function TargetClose(ByRef Pos As WorldPos) As Boolean

    '***************************************************
    'Author: ZaMa
    'Last Modification: 05/07/2010
    'Checks if there is any attackable user sorrounding the given position.
    '***************************************************
    Dim UserIndex   As Integer
    
    Dim nPos        As WorldPos

    Dim headingloop As Byte
    
    For headingloop = eHeading.SOUTH To eHeading.EAST
        
        nPos = Pos
        Call HeadtoPos(headingloop, nPos)
        
        UserIndex = MapData(nPos.Map, nPos.X, nPos.Y).UserIndex
        
        If UserIndex <> 0 Then
            If UserAtacable(UserIndex) Then
                TargetClose = True
                Exit Function

            End If

        End If

    Next headingloop
    
End Function

Private Function CheckNearUserPets(ByVal NPCIndex As Integer, _
                                   ByVal UserIndex As Integer) As Integer

    '***************************************************
    'Author: ZaMa
    'Last Modification: 09/07/2010
    'Checks if there is a non-paralized user pet in cleric vision range.
    '***************************************************
    Dim PetCounter As Long

    Dim PetIndex   As Integer
    
    With UserList(UserIndex)
        
        If .NroMascotas = 0 Then Exit Function
        
        For PetCounter = 1 To MAXMASCOTAS
        
            PetIndex = .MascotasIndex(PetCounter)

            If PetIndex > 0 Then
            
                With Npclist(PetIndex)
                    
                    ' If it is visible for the Pretorian
                    If InVisionRangeNpc(NPCIndex, .Pos.X, .Pos.Y) Then
                        
                        ' If not paralized then it's a suitable target
                        If .flags.Paralizado = 0 Then
                            CheckNearUserPets = PetIndex
                            Exit Function

                        End If
                    
                    End If
                    
                End With
               
            End If
            
        Next PetCounter
        
    End With

End Function

Private Function EsClaseMagica(ByVal UserIndex As Integer) As Boolean

    '***************************************************
    'Author: Unkown
    'Last Modification: -
    '***************************************************
    On Error GoTo ErrHandler

    With UserList(UserIndex)
        EsClaseMagica = .clase = eClass.Mage Or .clase = eClass.Cleric Or .clase = eClass.Druid Or .clase = eClass.Bard

    End With
    
    Exit Function

ErrHandler:
    LogError ("Error en EsClaseMagica. Error: " & Err.Number & " - " & Err.description)

End Function

Private Sub GreedyWalkTo(ByVal NPCIndex As Integer, _
                         ByVal TargetMap As Integer, _
                         ByVal TargetX As Integer, _
                         ByVal TargetY As Integer)

    '***************************************************
    'Author: Unknown
    '  Este procedimiento es llamado cada vez que un NPC deba ir
    '  a otro lugar en el mismo TargetMap. Utiliza una tecnica
    '  de programacion greedy no deterministica.
    '  Cada paso azaroso que me acerque al destino, es un buen paso.
    '  Si no hay mejor paso valido, entonces hay que volver atras y reintentar.
    '  Si no puedo moverme, me considero piketeado
    '  La funcion es larga, pero es O(1) - orden algoritmico temporal constante
    'Last Modification: 26/09/2010
    'Rapsodius - Changed Mod by And for speed
    '26/09/2010: ZaMa - Now make movements as normal npcs, which allows to kick caspers and invisible admins.
    '***************************************************
    On Error GoTo ErrHandler

    Dim NpcX      As Integer

    Dim NpcY      As Integer

    Dim RandomDir As Integer
    
    With Npclist(NPCIndex).Pos

        If .Map <> TargetMap Then Exit Sub
        
        NpcX = .X
        NpcY = .Y

    End With
    
    ' Arrived to destination
    If (NpcX = TargetX And NpcY = TargetY) Then Exit Sub
    
    ' Try to reach target
    If (NpcX > TargetX) Then
    
        ' Target is Down-Left
        If (NpcY < TargetY) Then
            
            RandomDir = RandomNumber(0, 9)

            If ((RandomDir And 1) = 0) Then ''move down
            
                If MoveNPCChar(NPCIndex, eHeading.SOUTH) Then
                    Exit Sub
                ElseIf MoveNPCChar(NPCIndex, eHeading.WEST) Then
                    Exit Sub
                ElseIf MoveNPCChar(NPCIndex, eHeading.EAST) Then
                    Exit Sub
                ElseIf MoveNPCChar(NPCIndex, eHeading.NORTH) Then
                    Exit Sub

                End If
                
            Else

                ''random first move
                If MoveNPCChar(NPCIndex, eHeading.WEST) Then
                    Exit Sub
                ElseIf MoveNPCChar(NPCIndex, eHeading.SOUTH) Then
                    Exit Sub
                ElseIf MoveNPCChar(NPCIndex, eHeading.EAST) Then
                    Exit Sub
                ElseIf MoveNPCChar(NPCIndex, eHeading.NORTH) Then
                    Exit Sub

                End If
                
            End If
            
            ' Target is Up-Left
        ElseIf (NpcY > TargetY) Then
        
            RandomDir = RandomNumber(0, 9)

            If ((RandomDir And 1) = 0) Then ''move up
            
                If MoveNPCChar(NPCIndex, eHeading.NORTH) Then
                    Exit Sub
                ElseIf MoveNPCChar(NPCIndex, eHeading.WEST) Then
                    Exit Sub
                ElseIf MoveNPCChar(NPCIndex, eHeading.SOUTH) Then
                    Exit Sub
                ElseIf MoveNPCChar(NPCIndex, eHeading.EAST) Then
                    Exit Sub

                End If
                
            Else

                ''random first move
                If MoveNPCChar(NPCIndex, eHeading.WEST) Then
                    Exit Sub
                ElseIf MoveNPCChar(NPCIndex, eHeading.NORTH) Then
                    Exit Sub
                ElseIf MoveNPCChar(NPCIndex, eHeading.SOUTH) Then
                    Exit Sub
                ElseIf MoveNPCChar(NPCIndex, eHeading.EAST) Then
                    Exit Sub

                End If
                
            End If
            
            ' Target is Straight Left
        Else
        
            If MoveNPCChar(NPCIndex, eHeading.WEST) Then
                Exit Sub
            ElseIf MoveNPCChar(NPCIndex, eHeading.SOUTH) Then
                Exit Sub
            ElseIf MoveNPCChar(NPCIndex, eHeading.NORTH) Then
                Exit Sub
            Else
                ' If moves to east, algorithm fails because starts a loop
                Call MoveFailed(NPCIndex)

            End If
            
        End If
    
    ElseIf (NpcX < TargetX) Then
        
        ' Target is Down-Right
        If (NpcY < TargetY) Then
            
            RandomDir = RandomNumber(0, 9)

            If ((RandomDir And 1) = 0) Then ''move down
            
                If MoveNPCChar(NPCIndex, eHeading.SOUTH) Then
                    Exit Sub
                ElseIf MoveNPCChar(NPCIndex, eHeading.EAST) Then
                    Exit Sub
                ElseIf MoveNPCChar(NPCIndex, eHeading.WEST) Then
                    Exit Sub
                ElseIf MoveNPCChar(NPCIndex, eHeading.NORTH) Then
                    Exit Sub

                End If
                
            Else    ''random first move
                
                If MoveNPCChar(NPCIndex, eHeading.EAST) Then
                    Exit Sub
                ElseIf MoveNPCChar(NPCIndex, eHeading.SOUTH) Then
                    Exit Sub
                ElseIf MoveNPCChar(NPCIndex, eHeading.WEST) Then
                    Exit Sub
                ElseIf MoveNPCChar(NPCIndex, eHeading.NORTH) Then
                    Exit Sub

                End If
                
            End If
        
            ' Target is Up-Right
        ElseIf (NpcY > TargetY) Then
        
            RandomDir = RandomNumber(0, 9)

            If ((RandomDir And 1) = 0) Then ''move up
            
                If MoveNPCChar(NPCIndex, eHeading.NORTH) Then
                    Exit Sub
                ElseIf MoveNPCChar(NPCIndex, eHeading.EAST) Then
                    Exit Sub
                ElseIf MoveNPCChar(NPCIndex, eHeading.WEST) Then
                    Exit Sub
                ElseIf MoveNPCChar(NPCIndex, eHeading.SOUTH) Then
                    Exit Sub

                End If
                
            Else
            
                If MoveNPCChar(NPCIndex, eHeading.EAST) Then
                    Exit Sub
                ElseIf MoveNPCChar(NPCIndex, eHeading.NORTH) Then
                    Exit Sub
                ElseIf MoveNPCChar(NPCIndex, eHeading.SOUTH) Then
                    Exit Sub
                ElseIf MoveNPCChar(NPCIndex, eHeading.WEST) Then
                    Exit Sub

                End If
                
            End If
        
            ' Target is Straight Right
        Else
        
            If MoveNPCChar(NPCIndex, eHeading.EAST) Then
                Exit Sub
            ElseIf MoveNPCChar(NPCIndex, eHeading.SOUTH) Then
                Exit Sub
            ElseIf MoveNPCChar(NPCIndex, eHeading.NORTH) Then
                Exit Sub
            Else
                ' If moves to west, algorithm fails because starts a loop
                Call MoveFailed(NPCIndex)

            End If
            
        End If
    
        ' Target is straight Up/Down
    Else
    
        ' Target is straight Up
        If (NpcY > TargetY) Then
        
            If MoveNPCChar(NPCIndex, eHeading.NORTH) Then
                Exit Sub
            ElseIf MoveNPCChar(NPCIndex, eHeading.EAST) Then
                Exit Sub
            ElseIf MoveNPCChar(NPCIndex, eHeading.WEST) Then
                Exit Sub
            Else
                ' If moves to south, algorithm fails because starts a loop
                Call MoveFailed(NPCIndex)

            End If
            
            ' Target is straight Down
        Else
        
            If MoveNPCChar(NPCIndex, eHeading.SOUTH) Then
                Exit Sub
            ElseIf MoveNPCChar(NPCIndex, eHeading.EAST) Then
                Exit Sub
            ElseIf MoveNPCChar(NPCIndex, eHeading.WEST) Then
                Exit Sub
            Else
                ' If moves to north, algorithm fails because starts a loop
                Call MoveFailed(NPCIndex)

            End If
            
        End If
        
    End If
    
    Exit Sub

ErrHandler:
    LogError ("Error en GreedyWalkTo. Error: " & Err.Number & " - " & Err.description)

End Sub

Private Sub MoveFailed(ByVal NPCIndex As Integer)
    '***************************************************
    'Author: ZaMa
    'Last Modification: 26/06/2010
    'Npc shouts a message overhead because cannot move (to avoid a loop)
    '***************************************************

    With Npclist(NPCIndex)

        If IntervaloPermiteAtacarNpc(NPCIndex) Then
            If (RandomNumber(1, 100) > 95) Then
                Call SendData(SendTarget.ToNPCArea, NPCIndex, PrepareMessageChatOverHead("Maldito bastardo, Ven aqui!", .Char.CharIndex, 255, 255, 0))

            End If

        End If

    End With

End Sub

Private Sub ReturnToCenter(ByVal NPCIndex As Integer)
    '***************************************************
    'Author: ZaMa
    'Last Modification: 19/09/2010
    'Returns to the center, where the king is.
    '***************************************************
    
    Call GreedyWalkTo(NPCIndex, CenterPos.Map, CenterPos.X, CenterPos.Y)
    
End Sub

Private Function FarFromTeam(ByVal NPCIndex) As Boolean
    '***************************************************
    'Author: Unknown & ZaMa
    'Last Modification: 19/09/2010
    'Checks if Npc is far from cleric protection ring
    '***************************************************
    
    If Npclist(NPCIndex).Pos.Map <> CenterPos.Map Then Exit Function
    
    FarFromTeam = (Abs(Npclist(NPCIndex).Pos.Y - CenterPos.Y) > MAX_DISTANCE) Or (Abs(Npclist(NPCIndex).Pos.X - CenterPos.X) > MAX_DISTANCE)
    
    Exit Function

ErrHandler:
    LogError ("Error en FarFromTeam. Error: " & Err.Number & " - " & Err.description)

End Function

Private Function UserReachable(ByVal NPCIndex As Integer, _
                               ByVal UserIndex As Integer) As Boolean

    '***************************************************
    'Author: ZaMa
    'Last Modification: 19/09/2010
    'Ignores users who are too far in order to avoid being separated from the rest of the team.
    '***************************************************
    On Error GoTo ErrHandler
        
    If Npclist(NPCIndex).Pos.Map <> CenterPos.Map Then Exit Function

    UserReachable = (Abs(UserList(UserIndex).Pos.X - CenterPos.X) < MAX_DISTANCE) And (Abs(UserList(UserIndex).Pos.Y - CenterPos.Y) < MAX_DISTANCE)
    
    Exit Function
    
ErrHandler:
    LogError ("Error en UserReachable. Error: " & Err.Number & " - " & Err.description)

End Function

Private Function InVisionRange(ByVal UserIndex As Integer, _
                               ByVal X As Integer, _
                               ByVal Y As Integer, _
                               Optional ByVal ExtendedRange As Boolean = False) As Boolean
    '***************************************************
    'Author: ZaMa
    'Last Modification: 24/06/2010
    'Checks whether user is in vision range or not
    '***************************************************

    Dim Rango As Integer

    Rango = val(IIf(ExtendedRange, 2, 1))
    
    With UserList(UserIndex)
        InVisionRange = Abs(.Pos.X - X) <= RANGO_VISION_X * Rango And Abs(.Pos.Y - Y) <= RANGO_VISION_Y * Rango

    End With
    
End Function

Private Function InVisionRangeNpc(ByVal NPCIndex As Integer, _
                                  ByVal X As Integer, _
                                  ByVal Y As Integer, _
                                  Optional ByVal ExtendedRange As Boolean = False) As Boolean
    '***************************************************
    'Author: ZaMa
    'Last Modification: 24/06/2010
    'Checks whether npc is in vision range or not
    '***************************************************
    
    Dim Rango As Integer

    Rango = val(IIf(ExtendedRange, 2, 1))

    With Npclist(NPCIndex)
        InVisionRangeNpc = Abs(.Pos.X - X) <= RANGO_VISION_NPC_X * Rango And Abs(.Pos.Y - Y) <= RANGO_VISION_NPC_Y * Rango

    End With
    
End Function

Private Function UserAtacable(ByVal UserIndex As Integer, _
                              Optional ByVal CheckVisibility As Boolean = True, _
                              Optional ByVal AttackAdmin As Boolean = True) As Boolean
    '***************************************************
    'Author: ZaMa
    'Last Modification: 05/10/2010
    'DEtermines whether the user can be atacked or not
    '05/10/2010: ZaMa - Now doesn't allow to attack admins sometimes.
    '***************************************************

    With UserList(UserIndex).flags
        UserAtacable = Not .EnConsulta And .AdminInvisible = 0 And .AdminPerseguible And .Muerto = 0
                       
        If CheckVisibility Then
            UserAtacable = UserAtacable And .Oculto = 0 And .invisible = 0

        End If
        
        If Not AttackAdmin Then
            UserAtacable = UserAtacable And (Not EsGm(UserIndex))

        End If
        
    End With
                        
End Function

Private Function UserDistance(ByVal UserIndex As Integer, _
                              ByVal X As Integer, _
                              ByVal Y As Integer) As Integer
    '***************************************************
    'Author: ZaMa
    'Last Modification: 24/06/2010
    'Calculates the user distance according to the given coords.
    '***************************************************

    With UserList(UserIndex)
        UserDistance = Abs(X - .Pos.X) + Abs(Y - .Pos.Y)

    End With
    
End Function

Private Function NpcDistance(ByVal NPCIndex As Integer, _
                             ByVal X As Integer, _
                             ByVal Y As Integer) As Integer
    '***************************************************
    'Author: ZaMa
    'Last Modification: 24/06/2010
    'Calculates the npc distance according to the given coords.
    '***************************************************

    With Npclist(NPCIndex)
        NpcDistance = Abs(X - .Pos.X) + Abs(Y - .Pos.Y)

    End With
    
End Function

Public Sub MuerePretoriano(ByVal NPCIndex As Integer)
    '***************************************************
    'Author: ZaMa
    'Last Modification: 27/06/2010
    'Eliminates the pretorian from the current alive ones, and respawn the clan if in pretorian's default map.
    '***************************************************

    Dim PretoIndex As Integer
    
    For PretoIndex = 1 To NRO_PRETORIANOS

        ' Remove npc index
        If Pretorianos(PretoIndex).NPCIndex = NPCIndex Then
            Pretorianos(PretoIndex).NPCIndex = 0
            ClanMembersAlive = ClanMembersAlive - 1
            
            ' Entire clan has been defeated
            If ClanMembersAlive = 0 Then

                ' Respawn it? (Only allowed in pretorian default map)
                If RespawnClan Then
                    
                    Dim NewSpawnPos As WorldPos
                
                    ' Switch respawn place (alternate between two places)
                    If CenterPos.X = LeftSpawnPos.X Then
                        NewSpawnPos = RightSpawnPos
                    Else
                        NewSpawnPos = LeftSpawnPos

                    End If
                    
                    ' Spawn it
                    Call SpawnClan(NewSpawnPos.Map, NewSpawnPos.X, NewSpawnPos.Y, ClanIndex, True)
                Else
                    ClanActive = False

                End If

            End If
            
            Exit Sub

        End If

    Next PretoIndex
    
End Sub

Private Function AllyParalyzed(ByVal NPCIndex As Integer, _
                               ByVal ExtendedRange As Boolean) As Integer
    '***************************************************
    'Author: ZaMa
    'Last Modification: 27/06/2010
    'Returns the index of the paralized ally if exists one.
    '***************************************************
   
    Dim PretoIndex As Integer

    Dim AllyIndex  As Integer
    
    For PretoIndex = 1 To NRO_PRETORIANOS
        
        AllyIndex = Pretorianos(PretoIndex).NPCIndex

        If AllyIndex <> 0 Then
            If Npclist(AllyIndex).flags.Paralizado = 1 Then
                If InVisionRangeNpc(AllyIndex, Npclist(NPCIndex).Pos.X, Npclist(NPCIndex).Pos.Y, ExtendedRange) Then
                    AllyParalyzed = AllyIndex
                    Exit Function

                End If

            End If

        End If
        
    Next PretoIndex
 
End Function

Private Function AllyInjured(ByVal NPCIndex As Integer) As Integer
    '***************************************************
    'Author: ZaMa
    'Last Modification: 09/07/2010
    'Returns the index of the first found injured ally if exists one.
    '***************************************************
   
    Dim PretoIndex As Integer

    Dim AllyIndex  As Integer
    
    For PretoIndex = 1 To NRO_PRETORIANOS
        
        AllyIndex = Pretorianos(PretoIndex).NPCIndex

        If AllyIndex <> 0 Then

            With Npclist(AllyIndex)

                If .Stats.MinHp < .Stats.MaxHp Then
                    If InVisionRangeNpc(NPCIndex, .Pos.X, .Pos.Y, False) Then
                        AllyInjured = AllyIndex
                        Exit Function

                    End If

                End If

            End With

        End If
        
    Next PretoIndex
 
End Function

Private Function AllyPoisoned(ByVal NPCIndex As Integer) As Integer
    '***************************************************
    'Author: ZaMa
    'Last Modification: 19/09/2010
    'Returns the index of the poisoned ally if exists one.
    '***************************************************
   
    Dim PretoIndex As Integer

    Dim AllyIndex  As Integer
    
    For PretoIndex = 1 To NRO_PRETORIANOS
        
        AllyIndex = Pretorianos(PretoIndex).NPCIndex

        If AllyIndex <> 0 Then
            If Npclist(AllyIndex).flags.Envenenado = 1 Then
                If InVisionRangeNpc(AllyIndex, Npclist(NPCIndex).Pos.X, Npclist(NPCIndex).Pos.Y) Then
                    AllyPoisoned = AllyIndex
                    Exit Function

                End If

            End If

        End If
        
    Next PretoIndex
 
End Function

Private Function AllyIncinera(ByVal NPCIndex As Integer) As Integer
    '***************************************************
    'Author: Lorwik
    'Last Modification: 05/09/2020
    'Returns the index of the incinera ally if exists one.
    '***************************************************
   
    Dim PretoIndex As Integer

    Dim AllyIndex  As Integer
    
    For PretoIndex = 1 To NRO_PRETORIANOS
        
        AllyIndex = Pretorianos(PretoIndex).NPCIndex

        If AllyIndex <> 0 Then
            If Npclist(AllyIndex).flags.Incinerado = 1 Then
                If InVisionRangeNpc(AllyIndex, Npclist(NPCIndex).Pos.X, Npclist(NPCIndex).Pos.Y) Then
                    AllyIncinera = AllyIndex
                    Exit Function

                End If

            End If

        End If
        
    Next PretoIndex
 
End Function

Private Function AllyDead() As Integer
    '***************************************************
    'Author: ZaMa
    'Last Modification: 19/09/2010
    'Returns the Index of the first dead member found, if exists.
    '***************************************************

    Dim PretoIndex As Integer

    Dim AllyIndex  As Integer
    
    For PretoIndex = 1 To NRO_PRETORIANOS
        
        If Pretorianos(PretoIndex).NPCIndex = 0 Then
            AllyDead = PretoIndex
            Exit Function

        End If
        
    Next PretoIndex

End Function

Public Function CanAtackMember(ByVal NPCIndex As Integer) As Boolean
    '***************************************************
    'Author: ZaMa
    'Last Modification: 26/09/2010
    'Returns true if given NpcIndex belongs to a normal clan member, o if king's alone.
    '***************************************************
    
    ' King?
    If Pretorianos(1).NPCIndex = NPCIndex Then

        ' any member left? => Can't attack
        If ClanMembersAlive <> 1 Then Exit Function

    End If
    
    CanAtackMember = True
    
End Function

Public Property Get ClanMap() As Integer
    '***************************************************
    'Author: ZaMa
    'Last Modification: 29/10/2010
    'Returns the clan map
    '***************************************************
    ClanMap = CenterPos.Map

End Property

Public Property Get Active() As Boolean
    '***************************************************
    'Author: ZaMa
    'Last Modification: 29/10/2010
    'Returns true if clan is active.
    '***************************************************
    Active = ClanActive

End Property

Public Sub DeleteClan()
    '***************************************************
    'Author: ZaMa
    'Last Modification: 29/10/2010
    'Kill all alive members.
    '***************************************************

    Dim PretoIndex As Long

    Dim NPCIndex   As Integer
    
    For PretoIndex = 1 To NRO_PRETORIANOS
        
        NPCIndex = Pretorianos(PretoIndex).NPCIndex

        If NPCIndex <> 0 Then Call QuitarNPC(NPCIndex)
    
        Pretorianos(PretoIndex).NPCIndex = 0
        
    Next PretoIndex
    
    ClanActive = False
    
End Sub
