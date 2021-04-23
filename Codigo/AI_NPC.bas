Attribute VB_Name = "AI"
'Argentum Online 0.12.2
'Copyright (C) 2002 M?rquez Pablo Ignacio
'
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
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 n?mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C?digo Postal 1900
'Pablo Ignacio M?rquez

Option Explicit

Public Enum TipoAI

    ESTATICO = 1
    MueveAlAzar = 2
    NpcMaloAtacaUsersBuenos = 3
    NPCDEFENSA = 4
    GuardiasAtacanCriminales = 5
    NpcObjeto = 6
    SigueAmo = 8
    NpcAtacaNpc = 9
    NpcPathfinding = 10
    
    'Pretorianos
    SacerdotePretorianoAi = 20
    GuerreroPretorianoAi = 21
    MagoPretorianoAi = 22
    CazadorPretorianoAi = 23
    ReyPretoriano = 24

End Enum

Public Const ELEMENTALFUEGO  As Integer = 93
Public Const ELEMENTALTIERRA As Integer = 94
Public Const ELEMENTALAGUA   As Integer = 92

' WyroX: Tiles extra que ve el NPC mayor que los usuarios.
' Para que no queden tontos en un borde de la pantalla y puedas atacarlos.
Private Const VISION_EXTRA         As Byte = 2
Public Const RANGO_VISION_NPC_X    As Byte = RANGO_VISION_X + VISION_EXTRA
Public Const RANGO_VISION_NPC_Y    As Byte = RANGO_VISION_Y + VISION_EXTRA

Private Const MINI_RANGO_X         As Byte = 3
Private Const MINI_RANGO_Y         As Byte = 3

'????????????????????????????????????????????????????????
'????????????????????????????????????????????????????????
'????????????????????????????????????????????????????????
'                        Modulo AI_NPC
'????????????????????????????????????????????????????????
'????????????????????????????????????????????????????????
'????????????????????????????????????????????????????????
'AI de los NPC
'????????????????????????????????????????????????????????
'????????????????????????????????????????????????????????
'????????????????????????????????????????????????????????

Private Sub GuardiasAI(ByVal NPCIndex As Integer, ByVal DelCaos As Boolean)

    '***************************************************
    'Autor: Unknown (orginal version)
    'Last Modification: 12/01/2010 (ZaMa)
    '14/09/2009: ZaMa - Now npcs don't atack protected users.
    '12/01/2010: ZaMa - Los npcs no atacan druidas mimetizados con npcs
    '***************************************************
    Dim nPos          As WorldPos

    Dim headingloop   As Byte

    Dim UI            As Integer

    Dim UserProtected As Boolean
    
    With Npclist(NPCIndex)

        For headingloop = eHeading.SOUTH To eHeading.EAST
            nPos = .Pos

            If .flags.Inmovilizado = 0 Or headingloop = .Char.Heading Then
                Call HeadtoPos(headingloop, nPos)

                If InMapBounds(nPos.Map, nPos.X, nPos.Y) Then
                    UI = MapData(nPos.Map, nPos.X, nPos.Y).UserIndex

                    If UI > 0 Then
                        UserProtected = Not IntervaloPermiteSerAtacado(UI) And UserList(UI).flags.NoPuedeSerAtacado
                        UserProtected = UserProtected Or UserList(UI).flags.Ignorado Or UserList(UI).flags.EnConsulta
                        
                        If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.AdminPerseguible And Not UserProtected Then

                            '?ES CRIMINAL?
                            If Not DelCaos Then
                                If criminal(UI) Then
                                    If NpcAtacaUser(NPCIndex, UI) Then
                                        Call ChangeNPCChar(NPCIndex, .Char.body, .Char.Head, headingloop)

                                    End If

                                    Exit Sub
                                ElseIf .flags.AttackedBy = UserList(UI).Name And Not .flags.Follow Then
                                    
                                    If NpcAtacaUser(NPCIndex, UI) Then
                                        Call ChangeNPCChar(NPCIndex, .Char.body, .Char.Head, headingloop)

                                    End If

                                    Exit Sub

                                End If

                            Else

                                If Not criminal(UI) Then
                                    If NpcAtacaUser(NPCIndex, UI) Then
                                        Call ChangeNPCChar(NPCIndex, .Char.body, .Char.Head, headingloop)

                                    End If

                                    Exit Sub
                                ElseIf .flags.AttackedBy = UserList(UI).Name And Not .flags.Follow Then
                                      
                                    If NpcAtacaUser(NPCIndex, UI) Then
                                        Call ChangeNPCChar(NPCIndex, .Char.body, .Char.Head, headingloop)

                                    End If

                                    Exit Sub

                                End If

                            End If

                        End If

                    End If

                End If

            End If  'not inmovil

        Next headingloop

    End With
    
    Call RestoreOldMovement(NPCIndex)

End Sub

''
' Handles the evil npcs' artificial intelligency.
'
' @param NpcIndex Specifies reference to the npc
Private Sub HostilMalvadoAI(ByVal NPCIndex As Integer)

    '**************************************************************
    'Author: Unknown
    'Last Modify Date: 12/01/2010 (ZaMa)
    '28/04/2009: ZaMa - Now those NPCs who doble attack, have 50% of posibility of casting a spell on user.
    '14/09/200*: ZaMa - Now npcs don't atack protected users.
    '12/01/2010: ZaMa - Los npcs no atacan druidas mimetizados con npcs
    '**************************************************************
    Dim nPos          As WorldPos

    Dim headingloop   As Byte

    Dim UI            As Integer

    Dim NPCI          As Integer

    Dim atacoPJ       As Boolean

    Dim UserProtected As Boolean
    
    atacoPJ = False
    
    With Npclist(NPCIndex)

        For headingloop = eHeading.SOUTH To eHeading.EAST
            nPos = .Pos

            Call HeadtoPos(headingloop, nPos)

            If InMapBounds(nPos.Map, nPos.X, nPos.Y) Then
                UI = MapData(nPos.Map, nPos.X, nPos.Y).UserIndex
                NPCI = MapData(nPos.Map, nPos.X, nPos.Y).NPCIndex

                If UI > 0 And Not atacoPJ Then
                    UserProtected = Not IntervaloPermiteSerAtacado(UI) And UserList(UI).flags.NoPuedeSerAtacado
                    UserProtected = UserProtected Or UserList(UI).flags.Ignorado Or UserList(UI).flags.EnConsulta
                        
                    If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.AdminPerseguible And (Not UserProtected) Then
                            
                        atacoPJ = True

                        If .Movement = NpcObjeto Then

                            ' Los npc objeto no atacan siempre al mismo usuario
                            If RandomNumber(1, 3) = 3 Then atacoPJ = False

                        End If
                            
                        If atacoPJ Then
                            If .flags.LanzaSpells Then
                                If .flags.AtacaDoble Then
                                    If (RandomNumber(0, 1)) Then
                                        If NpcAtacaUser(NPCIndex, UI) Then
                                            Call ChangeNPCChar(NPCIndex, .Char.body, .Char.Head, headingloop)

                                        End If

                                        Exit Sub

                                    End If

                                End If
                                    
                                Call ChangeNPCChar(NPCIndex, .Char.body, .Char.Head, headingloop)
                                Call NpcLanzaUnSpell(NPCIndex, UI)

                            End If

                        End If

                        If NpcAtacaUser(NPCIndex, UI) Then
                            Call ChangeNPCChar(NPCIndex, .Char.body, .Char.Head, headingloop)

                        End If

                        Exit Sub

                    End If

                ElseIf NPCI > 0 Then

                    If Npclist(NPCI).MaestroUser > 0 And Npclist(NPCI).flags.Paralizado = 0 Then
                        Call ChangeNPCChar(NPCIndex, .Char.body, .Char.Head, headingloop)
                        Call SistemaCombate.NpcAtacaNpc(NPCIndex, NPCI, False)
                        Exit Sub

                    End If

                End If

                End If

        Next headingloop

    End With
    
    Call RestoreOldMovement(NPCIndex)

End Sub

Private Sub HostilBuenoAI(ByVal NPCIndex As Integer)

    '***************************************************
    'Autor: Unknown (orginal version)
    'Last Modification: 12/01/2010 (ZaMa)
    '14/09/2009: ZaMa - Now npcs don't atack protected users.
    '12/01/2010: ZaMa - Los npcs no atacan druidas mimetizados con npcs
    '***************************************************
    Dim nPos          As WorldPos

    Dim headingloop   As eHeading

    Dim UI            As Integer

    Dim UserProtected As Boolean
    
    With Npclist(NPCIndex)

        For headingloop = eHeading.SOUTH To eHeading.EAST
            nPos = .Pos

            If .flags.Inmovilizado = 0 Or .Char.Heading = headingloop Then
                Call HeadtoPos(headingloop, nPos)

                If InMapBounds(nPos.Map, nPos.X, nPos.Y) Then
                    UI = MapData(nPos.Map, nPos.X, nPos.Y).UserIndex

                    If UI > 0 Then
                        If UserList(UI).Name = .flags.AttackedBy Then
                        
                            UserProtected = Not IntervaloPermiteSerAtacado(UI) And UserList(UI).flags.NoPuedeSerAtacado
                            UserProtected = UserProtected Or UserList(UI).flags.Ignorado Or UserList(UI).flags.EnConsulta
                            
                            If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.AdminPerseguible And Not UserProtected Then
                                If .flags.LanzaSpells > 0 Then
                                    Call NpcLanzaUnSpell(NPCIndex, UI)

                                End If
                                
                                If NpcAtacaUser(NPCIndex, UI) Then
                                    Call ChangeNPCChar(NPCIndex, .Char.body, .Char.Head, headingloop)

                                End If

                                Exit Sub

                            End If

                        End If

                    End If

                End If

            End If

        Next headingloop

    End With
    
    Call RestoreOldMovement(NPCIndex)

End Sub

Private Sub IrUsuarioCercano(ByVal NPCIndex As Integer)

    '***************************************************
    'Autor: Unknown (orginal version)
    'Last Modification: 25/07/2010 (ZaMa)
    '14/09/2009: ZaMa - Now npcs don't follow protected users.
    '12/01/2010: ZaMa - Los npcs no atacan druidas mimetizados con npcs
    '25/07/2010: ZaMa - Agrego una validacion temporal para evitar que los npcs ataquen a usuarios de mapas difernetes.
    '***************************************************
    Dim tHeading      As Byte

    Dim UserIndex     As Integer

    Dim SignoNS       As Integer

    Dim SignoEO       As Integer

    Dim i             As Long

    Dim UserProtected As Boolean
    
    With Npclist(NPCIndex)

        '¿Esta inmovilizado?
        If .flags.Inmovilizado = 1 Then

            Select Case .Char.Heading

                Case eHeading.NORTH
                    SignoNS = -1
                    SignoEO = 0
                
                Case eHeading.EAST
                    SignoNS = 0
                    SignoEO = 1
                
                Case eHeading.SOUTH
                    SignoNS = 1
                    SignoEO = 0
                
                Case eHeading.WEST
                    SignoEO = -1
                    SignoNS = 0

            End Select
            
            For i = 1 To Areas.ConnGroups(.Pos.Map).Count()
                UserIndex = Areas.ConnGroups(.Pos.Map).Item(i)
                
                '¿Esta en el rango de vision?
                If Abs(UserList(UserIndex).Pos.X - .Pos.X) <= RANGO_VISION_NPC_X And Sgn(UserList(UserIndex).Pos.X - .Pos.X) = SignoEO Then
                    If Abs(UserList(UserIndex).Pos.Y - .Pos.Y) <= RANGO_VISION_NPC_Y And Sgn(UserList(UserIndex).Pos.Y - .Pos.Y) = SignoNS Then
                        
                        UserProtected = Not IntervaloPermiteSerAtacado(UserIndex) And UserList(UserIndex).flags.NoPuedeSerAtacado
                        UserProtected = UserProtected Or UserList(UserIndex).flags.Ignorado Or UserList(UserIndex).flags.EnConsulta
                        
                        If UserList(UserIndex).flags.Muerto = 0 Then
                            If Not UserProtected Then
                                If .flags.LanzaSpells <> 0 Then Call NpcLanzaUnSpell(NPCIndex, UserIndex)
                                Exit Sub

                            End If

                        End If
                        
                    End If

                End If

            Next i
            
            ' No esta inmobilizado
        Else
            
            '¿El NPC se salio de su zona de origen? ¡Tiene que volver!
            If MapData(.Pos.Map, .Pos.X, .Pos.Y).ZonaIndex <> .ZonaOrig Then
                Debug.Print "No estoy en mi zona de origen"
                .PFINFO.PathLenght = 0
                Call NPCVuelveOrigin(NPCIndex)
                Exit Sub

            End If

            ' Tiene prioridad de seguir al usuario al que le pertenece si esta en el rango de vision
            Dim OwnerIndex As Integer
            
            OwnerIndex = .Owner

            '¿Tiene propietario?
            If OwnerIndex > 0 Then
                
                ' TODO: Es temporal hatsa reparar un bug que hace que ataquen a usuarios de otros mapas
                If UserList(OwnerIndex).Pos.Map = .Pos.Map Then
                    
                    'Is it in it's range of vision??
                    If Abs(UserList(OwnerIndex).Pos.X - .Pos.X) <= RANGO_VISION_NPC_X Then
                        If Abs(UserList(OwnerIndex).Pos.Y - .Pos.Y) <= RANGO_VISION_NPC_Y Then
                            
                            ' va hacia el si o esta invi ni oculto
                            If UserList(OwnerIndex).flags.invisible = 0 And UserList(OwnerIndex).flags.Oculto = 0 And Not UserList(OwnerIndex).flags.EnConsulta And Not UserList(OwnerIndex).flags.Ignorado Then
                                If .flags.LanzaSpells <> 0 Then Call NpcLanzaUnSpell(NPCIndex, OwnerIndex)
                                        
                                If Not .PFINFO.PathLenght > 0 Then tHeading = FindDirection(.Pos, UserList(OwnerIndex).Pos)
                                
                                If tHeading = 0 Then
                                
                                    If ReCalculatePath(NPCIndex) Then
                                        Call PathFindingAI(NPCIndex)

                                        'Existe el camino?
                                        If .PFINFO.NoPath Then 'Si no existe nos movemos al azar
                                            'Move randomly
                                            Call MoveNPCChar(NPCIndex, RandomNumber(eHeading.SOUTH, eHeading.EAST))

                                        End If
                                         
                                    Else

                                        If Not PathEnd(NPCIndex) Then
                                            Call FollowPath(NPCIndex)
                                        Else
                                            .PFINFO.PathLenght = 0

                                        End If

                                    End If
                                     
                                Else

                                    If Not .PFINFO.PathLenght > 0 Then Call MoveNPCChar(NPCIndex, tHeading)
                                    Exit Sub
                                     
                                End If

                                Exit Sub
    
                            End If

                        End If
                        
                    End If
                
                    ' Esto significa que esta bugueado.. Lo logueo, y "reparo" el error a mano (Todo temporal)
                Else
                    Call LogError("El npc: " & .Name & "(" & NPCIndex & "), intenta atacar a " & UserList(OwnerIndex).Name & "(Index: " & OwnerIndex & ", Mapa: " & UserList(OwnerIndex).Pos.Map & ") desde el mapa " & .Pos.Map)
                    .Owner = 0

                End If
                
            End If
            
            '¿No tiene propietario? Buscamos al usuario mas cercano
            For i = 1 To Areas.ConnGroups(.Pos.Map).Count()
                UserIndex = Areas.ConnGroups(.Pos.Map).Item(i)
                
                '¿Esta en el rango de vision?
                If Abs(UserList(UserIndex).Pos.X - .Pos.X) <= RANGO_VISION_NPC_X Then
                    If Abs(UserList(UserIndex).Pos.Y - .Pos.Y) <= RANGO_VISION_NPC_Y Then
                      
                        With UserList(UserIndex)
                                
                            UserProtected = Not IntervaloPermiteSerAtacado(UserIndex) And .flags.NoPuedeSerAtacado
                            UserProtected = UserProtected Or .flags.Ignorado Or .flags.EnConsulta
                                
                            'Si el user no esta muerto, invisible, protegido, etc...
                            If .flags.Muerto = 0 And .flags.invisible = 0 And .flags.Oculto = 0 And .flags.AdminPerseguible And Not UserProtected Then
                                    
                                If Npclist(NPCIndex).flags.LanzaSpells <> 0 Then Call NpcLanzaUnSpell(NPCIndex, UserIndex)
                                    
                                If Not Npclist(NPCIndex).PFINFO.PathLenght > 0 Then tHeading = FindDirection(Npclist(NPCIndex).Pos, .Pos)
                                    
                                If tHeading = 0 Then
                                    Call PathFindingAI(NPCIndex)
    
                                    If Not ReCalculatePath(NPCIndex) Then
                                        If Not PathEnd(NPCIndex) Then
                                            Call FollowPath(NPCIndex)
                                        Else
                                            Npclist(NPCIndex).PFINFO.PathLenght = 0
    
                                        End If
    
                                    End If
    
                                Else
    
                                    If Not Npclist(NPCIndex).PFINFO.PathLenght > 0 Then Call MoveNPCChar(NPCIndex, tHeading)
                                    Exit Sub
    
                                End If
    
                                Exit Sub
    
                            End If
                                
                        End With
                            
                    End If

                End If

            Next i
            
            'Si no hay usuarios y no esta en su pos de respawn, hacemos que vuelva
            If .Pos.X <> .Orig.X Or .Pos.Y <> .Orig.Y Or Npclist(NPCIndex).ZonaOrig <> MapData(.Pos.Map, .Pos.X, .Pos.Y).ZonaIndex Then Call NPCVuelveOrigin(NPCIndex)
            
            'Si llega aca es que no hab?a ning?n usuario cercano vivo.
            'A bailar. Pablo (ToxicWaste)
            'If RandomNumber(0, 10) = 0 Then
            '    Call MoveNPCChar(NPCIndex, CByte(RandomNumber(eHeading.SOUTH, eHeading.EAST)))

            'End If
            
        End If

    End With
    
    Call RestoreOldMovement(NPCIndex)

End Sub

''
' Makes a Pet / Summoned Npc to Follow an enemy
'
' @param NpcIndex Specifies reference to the npc
Private Sub SeguirAgresor(ByVal NPCIndex As Integer)

    '**************************************************************
    'Author: Unknown
    'Last Modify by: Marco Vanotti (MarKoxX)
    'Last Modify Date: 08/16/2008
    '08/16/2008: MarKoxX - Now pets that do mel? attacks have to be near the enemy to attack.
    '**************************************************************
    Dim tHeading As Byte

    Dim UI       As Integer
    
    Dim i        As Long
    
    Dim SignoNS  As Integer

    Dim SignoEO  As Integer

    With Npclist(NPCIndex)

        If .flags.Paralizado = 1 Or .flags.Inmovilizado = 1 Then

            Select Case .Char.Heading

                Case eHeading.NORTH
                    SignoNS = -1
                    SignoEO = 0
                
                Case eHeading.EAST
                    SignoNS = 0
                    SignoEO = 1
                
                Case eHeading.SOUTH
                    SignoNS = 1
                    SignoEO = 0
                
                Case eHeading.WEST
                    SignoEO = -1
                    SignoNS = 0

            End Select

            For i = 1 To Areas.ConnGroups(.Pos.Map).Count()
                UI = Areas.ConnGroups(.Pos.Map).Item(i)

                'Is it in it's range of vision??
                If Abs(UserList(UI).Pos.X - .Pos.X) <= RANGO_VISION_NPC_X And Sgn(UserList(UI).Pos.X - .Pos.X) = SignoEO Then
                    If Abs(UserList(UI).Pos.Y - .Pos.Y) <= RANGO_VISION_NPC_Y And Sgn(UserList(UI).Pos.Y - .Pos.Y) = SignoNS Then

                        If UserList(UI).Name = .flags.AttackedBy Then
                            If .MaestroUser > 0 Then
                                If Not criminal(.MaestroUser) And Not criminal(UI) And (UserList(.MaestroUser).flags.Seguro Or UserList(.MaestroUser).Faccion.ArmadaReal = 1) Then
                                    Call WriteConsoleMsg(.MaestroUser, "La mascota no atacara a ciudadanos si eres miembro del ejercito real o tienes el seguro activado.", FontTypeNames.FONTTYPE_INFO)
                                    .flags.AttackedBy = vbNullString
                                    Exit Sub
    
                                End If
    
                            End If
    
                            If (UserList(UI).flags.Muerto = 0 And UserList(UI).flags.invisible = 0 And UserList(UI).flags.Oculto = 0) Or (.flags.SiguiendoGm = True) Then
                                If .flags.LanzaSpells > 0 Then
                                    Call NpcLanzaUnSpell(NPCIndex, UI)
                                Else
    
                                    If Distancia(UserList(UI).Pos, Npclist(NPCIndex).Pos) <= 1 Then
    
                                        ' TODO : Set this a separate AI for Elementals and Druid's pets
                                        If Npclist(NPCIndex).Numero <> 92 Then
                                            Call NpcAtacaUser(NPCIndex, UI)
    
                                        End If
    
                                    End If
    
                                End If
    
                                Exit Sub
    
                            End If
    
                        End If
                            
                    End If

                End If
                
            Next i

        Else

            '¿El NPC se salio de su zona de origen? ¡Tiene que volver!
            If MapData(.Pos.Map, .Pos.X, .Pos.Y).ZonaIndex <> .ZonaOrig Then
                Call NPCVuelveOrigin(NPCIndex)
                .PFINFO.PathLenght = 0
                Exit Sub

            End If

            For i = 1 To Areas.ConnGroups(.Pos.Map).Count()
                UI = Areas.ConnGroups(.Pos.Map).Item(i)
                
                'Is it in it's range of vision??
                If Abs(UserList(UI).Pos.X - .Pos.X) <= RANGO_VISION_NPC_X Then
                    If Abs(UserList(UI).Pos.Y - .Pos.Y) <= RANGO_VISION_NPC_Y Then
                            
                        If UserList(UI).Name = .flags.AttackedBy Then
                            If .MaestroUser > 0 Then
                                If Not criminal(.MaestroUser) And Not criminal(UI) And (UserList(.MaestroUser).flags.Seguro Or UserList(.MaestroUser).Faccion.ArmadaReal = 1) Then
                                    Call WriteConsoleMsg(.MaestroUser, "La mascota no atacara a ciudadanos si eres miembro del ejercito real o tienes el seguro activado.", FontTypeNames.FONTTYPE_INFO)
                                    .flags.AttackedBy = vbNullString
                                    Call FollowAmo(NPCIndex)
                                    Exit Sub
    
                                End If
    
                            End If
                                
                            If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.invisible = 0 And UserList(UI).flags.Oculto = 0 Then
                                If .flags.LanzaSpells > 0 Then
                                    Call NpcLanzaUnSpell(NPCIndex, UI)
                                Else
    
                                    If Distancia(UserList(UI).Pos, Npclist(NPCIndex).Pos) <= 1 Then
    
                                        ' TODO : Set this a separate AI for Elementals and Druid's pets
                                        If Npclist(NPCIndex).Numero <> 92 Then
                                            Call NpcAtacaUser(NPCIndex, UI)
    
                                        End If
    
                                    End If
    
                                End If
                                     
                                tHeading = FindDirection(.Pos, UserList(UI).Pos)
                                Call MoveNPCChar(NPCIndex, tHeading)
                                     
                                Exit Sub
    
                            End If
    
                        End If
                        
                    End If

                End If
                
            Next i

        End If

        'Si no hay usuarios y no esta en su pos de respawn, hacemos que vuelva
        If .Pos.X <> .Orig.X Or .Pos.Y <> .Orig.Y Or Npclist(NPCIndex).ZonaOrig <> MapData(.Pos.Map, .Pos.X, .Pos.Y).ZonaIndex Then Call NPCVuelveOrigin(NPCIndex)

    End With
    
    Call RestoreOldMovement(NPCIndex)

End Sub

Private Sub RestoreOldMovement(ByVal NPCIndex As Integer)

    With Npclist(NPCIndex)

        If .MaestroUser = 0 Then
            .Movement = .flags.OldMovement
            .Hostile = .flags.OldHostil
            .flags.AttackedBy = vbNullString
            .flags.SiguiendoGm = False
        End If

    End With

End Sub

Private Sub PersigueCiudadano(ByVal NPCIndex As Integer)

    '***************************************************
    'Autor: Unknown (orginal version)
    'Last Modification: 12/01/2010 (ZaMa)
    '14/09/2009: ZaMa - Now npcs don't follow protected users.
    '12/01/2010: ZaMa - Los npcs no atacan druidas mimetizados con npcs.
    '***************************************************
    Dim UserIndex     As Integer

    Dim tHeading      As Byte

    Dim i             As Long

    Dim UserProtected As Boolean
    
    With Npclist(NPCIndex)

        For i = 1 To Areas.ConnGroups(.Pos.Map).Count()
            UserIndex = Areas.ConnGroups(.Pos.Map).Item(i)
                
            'Is it in it's range of vision??
            If Abs(UserList(UserIndex).Pos.X - .Pos.X) <= RANGO_VISION_NPC_X Then
                If Abs(UserList(UserIndex).Pos.Y - .Pos.Y) <= RANGO_VISION_NPC_Y Then
                    
                    If Not criminal(UserIndex) Then
                    
                        UserProtected = Not IntervaloPermiteSerAtacado(UserIndex) And UserList(UserIndex).flags.NoPuedeSerAtacado
                        UserProtected = UserProtected Or UserList(UserIndex).flags.Ignorado Or UserList(UserIndex).flags.EnConsulta
                        
                        If UserList(UserIndex).flags.Muerto = 0 And UserList(UserIndex).flags.invisible = 0 And UserList(UserIndex).flags.Oculto = 0 And UserList(UserIndex).flags.AdminPerseguible And Not UserProtected Then
                            
                            If .flags.LanzaSpells > 0 Then
                                Call NpcLanzaUnSpell(NPCIndex, UserIndex)

                            End If

                            tHeading = FindDirection(.Pos, UserList(UserIndex).Pos)
                            Call MoveNPCChar(NPCIndex, tHeading)
                            Exit Sub

                        End If

                    End If
                    
                End If

            End If
            
        Next i

    End With
    
    Call RestoreOldMovement(NPCIndex)

End Sub

Private Sub PersigueCriminal(ByVal NPCIndex As Integer)

    '***************************************************
    'Autor: Unknown (orginal version)
    'Last Modification: 12/01/2010 (ZaMa)
    '14/09/2009: ZaMa - Now npcs don't follow protected users.
    '12/01/2010: ZaMa - Los npcs no atacan druidas mimetizados con npcs.
    '***************************************************
    Dim UserIndex     As Integer

    Dim tHeading      As Byte

    Dim i             As Long

    Dim SignoNS       As Integer

    Dim SignoEO       As Integer

    Dim UserProtected As Boolean
    
    With Npclist(NPCIndex)

        If .flags.Inmovilizado = 1 Then

            Select Case .Char.Heading

                Case eHeading.NORTH
                    SignoNS = -1
                    SignoEO = 0
                
                Case eHeading.EAST
                    SignoNS = 0
                    SignoEO = 1
                
                Case eHeading.SOUTH
                    SignoNS = 1
                    SignoEO = 0
                
                Case eHeading.WEST
                    SignoEO = -1
                    SignoNS = 0

            End Select
            
            For i = 1 To Areas.ConnGroups(.Pos.Map).Count()
                UserIndex = Areas.ConnGroups(.Pos.Map).Item(i)
                
                'Is it in it's range of vision??
                If Abs(UserList(UserIndex).Pos.X - .Pos.X) <= RANGO_VISION_NPC_X And Sgn(UserList(UserIndex).Pos.X - .Pos.X) = SignoEO Then
                    If Abs(UserList(UserIndex).Pos.Y - .Pos.Y) <= RANGO_VISION_NPC_Y And Sgn(UserList(UserIndex).Pos.Y - .Pos.Y) = SignoNS Then
                        
                        If criminal(UserIndex) Then

                            With UserList(UserIndex)
                                 
                                UserProtected = Not IntervaloPermiteSerAtacado(UserIndex) And .flags.NoPuedeSerAtacado
                                UserProtected = UserProtected Or UserList(UserIndex).flags.Ignorado Or UserList(UserIndex).flags.EnConsulta
                                 
                                If .flags.Muerto = 0 And .flags.invisible = 0 And .flags.Oculto = 0 And .flags.AdminPerseguible And Not UserProtected Then
                                     
                                    If Npclist(NPCIndex).flags.LanzaSpells > 0 Then
                                        Call NpcLanzaUnSpell(NPCIndex, UserIndex)

                                    End If

                                    Exit Sub

                                End If

                            End With

                        End If
                        
                    End If

                End If

            Next i

        Else

            For i = 1 To Areas.ConnGroups(.Pos.Map).Count()
                UserIndex = Areas.ConnGroups(.Pos.Map).Item(i)
                
                'Is it in it's range of vision??
                If Abs(UserList(UserIndex).Pos.X - .Pos.X) <= RANGO_VISION_NPC_X Then
                    If Abs(UserList(UserIndex).Pos.Y - .Pos.Y) <= RANGO_VISION_NPC_Y Then
                        
                        If criminal(UserIndex) Then
                            
                            UserProtected = Not IntervaloPermiteSerAtacado(UserIndex) And UserList(UserIndex).flags.NoPuedeSerAtacado
                            UserProtected = UserProtected Or UserList(UserIndex).flags.Ignorado
                            
                            If UserList(UserIndex).flags.Muerto = 0 And UserList(UserIndex).flags.invisible = 0 And UserList(UserIndex).flags.Oculto = 0 And UserList(UserIndex).flags.AdminPerseguible And Not UserProtected Then

                                If .flags.LanzaSpells > 0 Then
                                    Call NpcLanzaUnSpell(NPCIndex, UserIndex)

                                End If

                                If .flags.Inmovilizado = 1 Then Exit Sub
                                tHeading = FindDirection(.Pos, UserList(UserIndex).Pos)
                                Call MoveNPCChar(NPCIndex, tHeading)
                                Exit Sub

                            End If

                        End If
                        
                    End If

                End If
                
            Next i

        End If

    End With
    
    Call RestoreOldMovement(NPCIndex)

End Sub

Private Sub SeguirAmo(ByVal NPCIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim tHeading As Byte

    Dim UI       As Integer
    
    With Npclist(NPCIndex)

        If .Target = 0 And .TargetNPC = 0 Then
            UI = .MaestroUser
            
            If UI > 0 Then

                'Is it in it's range of vision??
                If Abs(UserList(UI).Pos.X - .Pos.X) <= RANGO_VISION_NPC_X Then
                    If Abs(UserList(UI).Pos.Y - .Pos.Y) <= RANGO_VISION_NPC_Y Then
                        If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.invisible = 0 And UserList(UI).flags.Oculto = 0 And Distancia(.Pos, UserList(UI).Pos) > 3 Then
                            tHeading = FindDirection(.Pos, UserList(UI).Pos)
                            Call MoveNPCChar(NPCIndex, tHeading)
                            Exit Sub

                        End If

                    End If

                End If

            End If

        End If

    End With
    
    Call RestoreOldMovement(NPCIndex)

End Sub

Private Sub AiNpcAtacaNpc(ByVal NPCIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim tHeading As Byte

    Dim X        As Long

    Dim Y        As Long

    Dim NI       As Integer

    Dim bNoEsta  As Boolean
    
    Dim SignoNS  As Integer

    Dim SignoEO  As Integer
    
    With Npclist(NPCIndex)

        If .flags.Inmovilizado = 1 Then

            Select Case .Char.Heading

                Case eHeading.NORTH
                    SignoNS = -1
                    SignoEO = 0
                
                Case eHeading.EAST
                    SignoNS = 0
                    SignoEO = 1
                
                Case eHeading.SOUTH
                    SignoNS = 1
                    SignoEO = 0
                
                Case eHeading.WEST
                    SignoEO = -1
                    SignoNS = 0

            End Select
            
            For Y = .Pos.Y To .Pos.Y + SignoNS * RANGO_VISION_NPC_Y Step IIf(SignoNS = 0, 1, SignoNS)
                For X = .Pos.X To .Pos.X + SignoEO * RANGO_VISION_NPC_X Step IIf(SignoEO = 0, 1, SignoEO)

                    If X >= MinXBorder And X <= MaxXBorder And Y >= MinYBorder And Y <= MaxYBorder Then
                        NI = MapData(.Pos.Map, X, Y).NPCIndex

                        If NI > 0 Then
                            If .TargetNPC = NI Then
                                bNoEsta = True

                                If .Numero = ELEMENTALFUEGO Then
                                    Call NpcLanzaUnSpellSobreNpc(NPCIndex, NI)

                                    If Npclist(NI).NPCtype = DRAGON Then
                                        Call NpcLanzaUnSpellSobreNpc(NI, NPCIndex)

                                    End If

                                Else

                                    'aca verificamosss la distancia de ataque
                                    If Distancia(.Pos, Npclist(NI).Pos) <= 1 Then
                                        Call SistemaCombate.NpcAtacaNpc(NPCIndex, NI)

                                    End If

                                End If

                                Exit Sub

                            End If

                        End If

                    End If

                Next X
            Next Y

        Else

            For Y = .Pos.Y - RANGO_VISION_NPC_Y To .Pos.Y + RANGO_VISION_NPC_Y
                For X = .Pos.X - RANGO_VISION_NPC_Y To .Pos.X + RANGO_VISION_NPC_Y

                    If X >= MinXBorder And X <= MaxXBorder And Y >= MinYBorder And Y <= MaxYBorder Then
                        NI = MapData(.Pos.Map, X, Y).NPCIndex

                        If NI > 0 Then
                            If .TargetNPC = NI Then
                                bNoEsta = True

                                If .Numero = ELEMENTALFUEGO Then
                                    Call NpcLanzaUnSpellSobreNpc(NPCIndex, NI)

                                    If Npclist(NI).NPCtype = DRAGON Then
                                        Call NpcLanzaUnSpellSobreNpc(NI, NPCIndex)

                                    End If

                                Else

                                    'aca verificamosss la distancia de ataque
                                    If Distancia(.Pos, Npclist(NI).Pos) <= 1 Then
                                        Call SistemaCombate.NpcAtacaNpc(NPCIndex, NI)

                                    End If

                                End If

                                If .flags.Inmovilizado = 1 Then Exit Sub
                                If .TargetNPC = 0 Then Exit Sub
                                tHeading = FindDirection(.Pos, Npclist(MapData(.Pos.Map, X, Y).NPCIndex).Pos)
                                Call MoveNPCChar(NPCIndex, tHeading)
                                Exit Sub

                            End If

                        End If

                    End If

                Next X
            Next Y

        End If
        
        If Not bNoEsta Then
            If .MaestroUser > 0 Then
                Call FollowAmo(NPCIndex)
            Else
                .Movement = .flags.OldMovement
                .Hostile = .flags.OldHostil

            End If

        End If

    End With

End Sub

Public Sub AiNpcObjeto(ByVal NPCIndex As Integer)

    '***************************************************
    'Autor: ZaMa
    'Last Modification: 14/09/2009 (ZaMa)
    '14/09/2009: ZaMa - Now npcs don't follow protected users.
    '***************************************************
    Dim UserIndex     As Integer

    Dim i             As Long

    Dim UserProtected As Boolean
    
    With Npclist(NPCIndex)

        For i = 1 To Areas.ConnGroups(.Pos.Map).Count()
            UserIndex = Areas.ConnGroups(.Pos.Map).Item(i)
            
            'Is it in it's range of vision??
            If Abs(UserList(UserIndex).Pos.X - .Pos.X) <= RANGO_VISION_NPC_X Then
                If Abs(UserList(UserIndex).Pos.Y - .Pos.Y) <= RANGO_VISION_NPC_Y Then
                    
                    With UserList(UserIndex)
                        UserProtected = Not IntervaloPermiteSerAtacado(UserIndex) And .flags.NoPuedeSerAtacado
                        
                        If .flags.Muerto = 0 And .flags.invisible = 0 And .flags.Oculto = 0 And .flags.AdminPerseguible And Not UserProtected Then
                            
                            ' No quiero que ataque siempre al primero
                            If RandomNumber(1, 3) < 3 Then
                                If Npclist(NPCIndex).flags.LanzaSpells > 0 Then
                                    Call NpcLanzaUnSpell(NPCIndex, UserIndex)

                                End If
                            
                                Exit Sub

                            End If

                        End If

                    End With

                End If

            End If
            
        Next i

    End With

End Sub

Public Sub NPCVuelveOrigin(ByVal NPCIndex As Integer)
'*****************************************
'Author: Lorwik
'Last Modify Date: 20/04/2021
'Descripcion: El NPC vuelve a su lugar de respawn
 '*****************************************
    Dim tHeading As Byte
 
    With Npclist(NPCIndex)
 
    If Not .PFINFO.PathLenght > 0 Then tHeading = FindDirection(.Pos, .Orig)
                                
        If tHeading = 0 Then
            Call PathFindingAI(NPCIndex)
    
            If Not ReCalculatePath(NPCIndex) Then
                If Not PathEnd(NPCIndex) Then
                    Call FollowPath(NPCIndex)
                    
                Else
                    Npclist(NPCIndex).PFINFO.PathLenght = 0
    
                End If
    
            End If
    
        Else
    
            If Not .PFINFO.PathLenght > 0 Then Call MoveNPCChar(NPCIndex, tHeading)
            Exit Sub
    
        End If
    
    End With

    Exit Sub

End Sub

Sub NPCAI(ByVal NPCIndex As Integer)

    '**************************************************************
    'Author: Unknown
    'Last Modify by: ZaMa
    'Last Modify Date: 15/11/2009
    '08/16/2008: MarKoxX - Now pets that do mel? attacks have to be near the enemy to attack.
    '15/11/2009: ZaMa - Implementacion de npc objetos ai.
    '**************************************************************
    On Error GoTo ErrorHandler

    With Npclist(NPCIndex)
    
        'TODO: Hay que comprobar si el usuario esta en la zona de ORIGEN del Npc.
    
        '¿Hay usuarios en la zona?
        If MapZonas(.Pos.Map, NPCZonaId(NPCIndex)).NumUsers > 0 Then
    
            'Cada NPC tiene su propia velocidad
            If IntervaloNpcVelocidadVariable(NPCIndex) Then
        
                '<<<<<<<<<<< Ataques >>>>>>>>>>>>>>>>
                If .MaestroUser = 0 Then
    
                    'Busca a alguien para atacar
                    '?Es un guardia?
               
                    If .NPCtype = eNPCType.GuardiaReal Then  '¿Es un guardia?
                        Call GuardiasAI(NPCIndex, False)
                    
                    ElseIf .NPCtype = eNPCType.Guardiascaos Then  '¿Es un guardia rebelde?
                        Call GuardiasAI(NPCIndex, True)
                    
                    ElseIf .Hostile And .Stats.Alineacion <> 0 Then '¿Es un NPC Hostil?
                        Call HostilMalvadoAI(NPCIndex)
                    
                    ElseIf .Hostile And .Stats.Alineacion = 0 Then '¿Es un NPC NO Hostil?
                        Call HostilBuenoAI(NPCIndex)
    
                    End If
    
                Else
    
                    'Evitamos que ataque a su amo, a menos
                    'que el amo lo ataque.
                    'Call HostilBuenoAI(NpcIndex)
                End If
        
                '<<<<<<<<<<<Movimiento>>>>>>>>>>>>>>>>
                Select Case .Movement
    
                    Case TipoAI.MueveAlAzar
    
                        If .flags.Inmovilizado = 1 Then Exit Sub
                        
                        If .NPCtype = eNPCType.GuardiaReal Then
                            If RandomNumber(1, 12) = 3 Then
                                Call MoveNPCChar(NPCIndex, CByte(RandomNumber(eHeading.SOUTH, eHeading.EAST)))
    
                            End If
                        
                            Call PersigueCriminal(NPCIndex)
                        
                        ElseIf .NPCtype = eNPCType.Guardiascaos Then
    
                            If RandomNumber(1, 12) = 3 Then
                                Call MoveNPCChar(NPCIndex, CByte(RandomNumber(eHeading.SOUTH, eHeading.EAST)))
    
                            End If
                        
                            Call PersigueCiudadano(NPCIndex)
                        
                        Else
    
                            If RandomNumber(1, 12) = 3 Then
                                Call MoveNPCChar(NPCIndex, CByte(RandomNumber(eHeading.SOUTH, eHeading.EAST)))
    
                            End If
    
                        End If
                
                        'Va hacia el usuario cercano
                    Case TipoAI.NpcMaloAtacaUsersBuenos
                        Call IrUsuarioCercano(NPCIndex)
                
                        'Va hacia el usuario que lo ataco(FOLLOW)
                    Case TipoAI.NPCDEFENSA
                        Call SeguirAgresor(NPCIndex)
                
                        'Persigue criminales
                    Case TipoAI.GuardiasAtacanCriminales
                        Call PersigueCriminal(NPCIndex)
                
                    Case TipoAI.SigueAmo
    
                        If .flags.Inmovilizado = 1 Then Exit Sub
                        Call SeguirAmo(NPCIndex)
    
                        If RandomNumber(1, 12) = 3 Then
                            Call MoveNPCChar(NPCIndex, CByte(RandomNumber(eHeading.SOUTH, eHeading.EAST)))
    
                        End If
                
                    Case TipoAI.NpcAtacaNpc
                        Call AiNpcAtacaNpc(NPCIndex)
                    
                    Case TipoAI.NpcObjeto
                        Call AiNpcObjeto(NPCIndex)
                    
                    Case TipoAI.NpcPathfinding

                        If .flags.Inmovilizado = 1 Then Exit Sub
                        If ReCalculatePath(NPCIndex) Then
                            Call PathFindingAI(NPCIndex)
    
                            'Existe el camino?
                            If .PFINFO.NoPath Then 'Si no existe nos movemos al azar
                                'Move randomly
                                Call MoveNPCChar(NPCIndex, RandomNumber(eHeading.SOUTH, eHeading.EAST))
    
                            End If
    
                        Else
    
                            If Not PathEnd(NPCIndex) Then
                                Call FollowPath(NPCIndex)
                            Else
                                .PFINFO.PathLenght = 0
    
                            End If
    
                        End If
    
                End Select
        
            End If
            
            Exit Sub
       End If
       
        'Si no hay usuarios y no esta en su pos de respawn, hacemos que vuelva
        If Npclist(NPCIndex).Pos.X <> Npclist(NPCIndex).Orig.X Or Npclist(NPCIndex).Pos.Y <> Npclist(NPCIndex).Orig.Y Then
            If IntervaloNpcVelocidadVariable(NPCIndex) Then Call NPCVuelveOrigin(NPCIndex)
            
        End If

    End With

    Exit Sub

ErrorHandler:

    With Npclist(NPCIndex)
        Call LogError("Error en NPCAI. Error: " & Err.Number & " - " & Err.description & ". " & "Npc: " & .Name & ", Index: " & NPCIndex & ", MaestroUser: " & .MaestroUser & ", MaestroNpc: " & .MaestroNpc & ", Mapa: " & .Pos.Map & " x:" & .Pos.X & " y:" & .Pos.Y & "ZONA: " & NPCZonaId(NPCIndex) & " Mov:" & .Movement & " TargU:" & .Target & " TargN:" & .TargetNPC)

    End With
    
    Dim MiNPC As NPC

    MiNPC = Npclist(NPCIndex)
    Call QuitarNPC(NPCIndex)
    Call ReSpawnNpc(MiNPC)

End Sub

Function UserNear(ByVal NPCIndex As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    'Returns True if there is an user adjacent to the npc position.
    '***************************************************

    With Npclist(NPCIndex)
        UserNear = Not Int(Distance(.Pos.X, .Pos.Y, UserList(.PFINFO.TargetUser).Pos.X, UserList(.PFINFO.TargetUser).Pos.Y)) > 1

    End With

End Function

Function ReCalculatePath(ByVal NPCIndex As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    'Returns true if we have to seek a new path
    '***************************************************

    If Npclist(NPCIndex).PFINFO.PathLenght = 0 Then
        ReCalculatePath = True
    ElseIf Not UserNear(NPCIndex) And Npclist(NPCIndex).PFINFO.PathLenght = Npclist(NPCIndex).PFINFO.CurPos - 1 Then
        ReCalculatePath = True

    End If

End Function

Function PathEnd(ByVal NPCIndex As Integer) As Boolean
    '***************************************************
    'Author: Gulfas Morgolock
    'Last Modification: -
    'Returns if the npc has arrived to the end of its path
    '***************************************************
    PathEnd = Npclist(NPCIndex).PFINFO.CurPos = Npclist(NPCIndex).PFINFO.PathLenght

End Function

Function FollowPath(ByVal NPCIndex As Integer) As Boolean

    '***************************************************
    'Author: Gulfas Morgolock
    'Last Modification: -
    'Moves the npc.
    '***************************************************
    Dim tmpPos   As WorldPos

    Dim tHeading As Byte
    
    With Npclist(NPCIndex)
        tmpPos.Map = .Pos.Map
        tmpPos.X = .PFINFO.Path(.PFINFO.CurPos).Y ' invert? las coordenadas
        tmpPos.Y = .PFINFO.Path(.PFINFO.CurPos).X
        
        'Debug.Print "(" & tmpPos.X & "," & tmpPos.Y & ")"
        
        tHeading = FindDirection(.Pos, tmpPos)
        
        MoveNPCChar NPCIndex, tHeading
        
        .PFINFO.CurPos = .PFINFO.CurPos + 1

    End With

End Function

Function PathFindingAI(ByVal NPCIndex As Integer) As Boolean

    '***************************************************
    'Author: Gulfas Morgolock
    'Last Modification: -
    'This function seeks the shortest path from the Npc
    'to the user's location.
    '***************************************************
    Dim Y           As Long
    Dim X           As Long
    
    With Npclist(NPCIndex)

        For Y = .Pos.Y - 10 To .Pos.Y + 10    'Makes a loop that looks at
            For X = .Pos.X - 10 To .Pos.X + 10   '5 tiles in every direction
                
                'Makñe sure tile is legal
                If X > MinXBorder And X < MaxXBorder And Y > MinYBorder And Y < MaxYBorder Then
                    
                    'look for a user
                    If MapData(.Pos.Map, X, Y).UserIndex > 0 Then

                        'Move towards user
                        Dim tmpUserIndex As Integer

                        tmpUserIndex = MapData(.Pos.Map, X, Y).UserIndex

                        With UserList(tmpUserIndex)

                            If .flags.Muerto = 0 And .flags.invisible = 0 And .flags.Oculto = 0 And .flags.AdminPerseguible Then
                            
                                'We have to invert the coordinates, this is because
                                'ORE refers to maps in converse way of my pathfinding
                                'routines.
                                Npclist(NPCIndex).PFINFO.Target.X = .Pos.Y
                                Npclist(NPCIndex).PFINFO.Target.Y = .Pos.X 'ops!
                                Npclist(NPCIndex).PFINFO.TargetUser = tmpUserIndex
                                Call SeekPath(NPCIndex)
                                
                                Exit Function

                            End If

                        End With

                    End If

                End If

            Next X
        Next Y

    End With

End Function

Sub NpcLanzaUnSpell(ByVal NPCIndex As Integer, ByVal UserIndex As Integer)

    '**************************************************************
    'Author: Unknown
    'Last Modify by: -
    'Last Modify Date: -
    '**************************************************************
    With UserList(UserIndex)

        If .flags.invisible = 1 Or .flags.Oculto = 1 Then Exit Sub

    End With
    
    Dim K As Integer

    K = RandomNumber(1, Npclist(NPCIndex).flags.LanzaSpells)
    Call NpcLanzaSpellSobreUser(NPCIndex, UserIndex, Npclist(NPCIndex).Spells(K))

End Sub

Sub NpcLanzaUnSpellSobreNpc(ByVal NPCIndex As Integer, ByVal TargetNPC As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim K As Integer

    K = RandomNumber(1, Npclist(NPCIndex).flags.LanzaSpells)
    Call NpcLanzaSpellSobreNpc(NPCIndex, TargetNPC, Npclist(NPCIndex).Spells(K))

End Sub

Public Sub SacerdoteHealUser(ByVal UserIndex As Integer)

    With UserList(UserIndex)

        'Si ya esta a full, no hacemos nada
        If .Stats.MinHp = .Stats.MaxHp And .Stats.MinMAN = .Stats.MaxMAN And .flags.Maldicion = 0 And .flags.Ceguera = 0 And .flags.Envenenado = 0 And .flags.Incinerado = 0 Then
            Call WriteChatOverHead(UserIndex, "Hijo mio, los dioses ya han sanado todas tus heridas.", Npclist(.flags.TargetNPC).Char.CharIndex, vbWhite)
            Exit Sub
        End If

        'Enviamos sonido de curar
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_CURAR_SACERDOTE, .Pos.X, .Pos.Y))

        .Stats.MinHp = .Stats.MaxHp

        Call WriteUpdateHP(UserIndex)

        Call WriteConsoleMsg(UserIndex, "El sacerdote te ha curado!!", FontTypeNames.FONTTYPE_INFO)

        Call SacerdoteHealEffectsAndRestoreMana(UserIndex)

        Call WriteUpdateUserStats(UserIndex)
    End With

End Sub

Public Sub SacerdoteResucitateUser(ByVal UserIndex As Integer)
    With UserList(UserIndex)

        'Enviamos sonido de resucitacion (Recox)
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_RESUCITAR_SACERDOTE, .Pos.X, .Pos.Y))
        
        Call RevivirUsuario(UserIndex)
        Call WriteConsoleMsg(UserIndex, "Has sido resucitado!!", FontTypeNames.FONTTYPE_INFO)

        'Si es newbie le sacamos todo, sino solo lo revivimos. (Recox)
        If EsNewbie(UserIndex) Then _
            Call SacerdoteHealEffectsAndRestoreMana(UserIndex)

    End With
End Sub

Private Sub SacerdoteHealEffectsAndRestoreMana(ByVal UserIndex As Integer)

    With UserList(UserIndex)
        ' Sacamos la maldicion.
        If .flags.Maldicion = 1 Then
            .flags.Maldicion = 0
            Call WriteConsoleMsg(UserIndex, "El sacerdote te ha curado de la maldicion.", FontTypeNames.FONTTYPE_INFO)
        End If
 
        ' Sacamos la ceguera.
        If .flags.Ceguera = 1 Then
            .flags.Ceguera = 0
            Call WriteConsoleMsg(UserIndex, "El sacerdote te ha curado de la ceguera.", FontTypeNames.FONTTYPE_INFO)
        End If

        ' Curamos su envenenamiento.
        If .flags.Envenenado = 1 Then
            .flags.Envenenado = 0
            Call WriteConsoleMsg(UserIndex, "El sacerdote te ha curado del envenenamiento.", FontTypeNames.FONTTYPE_INFO)
        End If
        
        ' Curamos su incineramiento
        If .flags.Incinerado = 1 Then
            .flags.Incinerado = 0
            Call WriteConsoleMsg(UserIndex, "El sacerdote apago tus llamas.", FontTypeNames.FONTTYPE_INFO)
        End If

        ' Restauramos su mana.
        .Stats.MinMAN = .Stats.MaxMAN
        Call WriteUpdateMana(UserIndex)
        Call WriteConsoleMsg(UserIndex, "El sacerdote te ha restaurado el mana completamente.", FontTypeNames.FONTTYPE_INFO)
    End With
End Sub

