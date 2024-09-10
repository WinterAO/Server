Attribute VB_Name = "modSendData"
'**************************************************************
' SendData.bas - Has all methods to send data to different user groups.
' Makes use of the modAreas module.
'
' Implemented by Juan Martin Sotuyo Dodero (Maraxus) (juansotuyo@gmail.com)
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

''
' Contains all methods to send data to different user groups.
' Makes use of the modAreas module.
'
' @author Juan Martin Sotuyo Dodero (Maraxus) juansotuyo@gmail.com
' @version 1.0.0
' @date 20070107

Option Explicit

Public Enum SendTarget

    Toall = 1
    ToUser
    toMap
    ToPCArea
    ToAllButIndex
    ToGM
    ToNPCArea
    ToGuildMembers
    ToAdmins
    ToPCAreaButIndex
    ToAdminsAreaButConsejeros
    ToDiosesYclan
    ToConsejo
    ToClanArea
    ToConsejoCaos
    ToRolesMasters
    ToDeadArea
    ToCiudadanos
    ToCriminales
    ToPartyArea
    ToReal
    ToCaos
    ToCiudadanosYRMs
    ToCriminalesYRMs
    ToRealYRMs
    ToCaosYRMs
    ToHigherAdmins
    ToGMsAreaButRmsOrCounselors
    ToUsersAreaButGMs
    ToUsersAndRmsAndCounselorsAreaButGMs

End Enum

Public Sub SendData(ByVal sndRoute As SendTarget, _
                    ByVal sndIndex As Integer, _
                    ByVal sndData As String, _
                    Optional ByVal IsDenounce As Boolean = False)

    '**************************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus) - Rewrite of original
    'Last Modify Date: 14/11/2010
    'Last modified by: ZaMa
    '14/11/2010: ZaMa - Now denounces can be desactivated.
    '**************************************************************
    On Error Resume Next

    Dim LoopC As Long
    
    Debug.Print "SendData> " & sndRoute

    Select Case sndRoute
    
        Case SendTarget.ToUser
            If UserList(sndIndex).ConnID <> -1 Then
                Call UserList(sndIndex).outgoingData.WriteASCIIStringFixed(sndData)

            End If

        Case SendTarget.ToPCArea
            Call SendToUserArea(sndIndex, sndData)
            Exit Sub
        
        Case SendTarget.ToAdmins

            For LoopC = 1 To LastUser

                If UserList(LoopC).ConnID <> -1 Then
                    If UserList(LoopC).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero) Then

                        ' Denounces can be desactivated
                        If IsDenounce Then
                            
                            If UserList(LoopC).flags.SendDenounces Then
                                Call UserList(LoopC).outgoingData.WriteASCIIStringFixed(sndData)
                            End If

                        Else
                            Call UserList(LoopC).outgoingData.WriteASCIIStringFixed(sndData)

                        End If

                    End If

                End If

            Next LoopC

            Exit Sub
        
        Case SendTarget.Toall

            For LoopC = 1 To LastUser

                If UserList(LoopC).ConnID <> -1 Then
                    
                    If UserList(LoopC).flags.UserLogged Then 'Esta logeado como usuario?
                        Call UserList(LoopC).outgoingData.WriteASCIIStringFixed(sndData)
                    End If

                End If

            Next LoopC

            Exit Sub
        
        Case SendTarget.ToAllButIndex

            For LoopC = 1 To LastUser

                If (UserList(LoopC).ConnID <> -1) And (LoopC <> sndIndex) Then
                    
                    If UserList(LoopC).flags.UserLogged Then 'Esta logeado como usuario?
                        Call UserList(LoopC).outgoingData.WriteASCIIStringFixed(sndData)
                    End If

                End If

            Next LoopC

            Exit Sub
        
        Case SendTarget.toMap
            Call SendToMap(sndIndex, sndData)
            Exit Sub
        
        Case SendTarget.ToGuildMembers
            LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)

            While LoopC > 0

                If (UserList(LoopC).ConnID <> -1) Then
                    Call UserList(LoopC).outgoingData.WriteASCIIStringFixed(sndData)
                End If

                LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
            Wend
            Exit Sub
        
        Case SendTarget.ToDeadArea
            Call SendToDeadUserArea(sndIndex, sndData)
            Exit Sub
        
        Case SendTarget.ToPCAreaButIndex
            Call SendToUserAreaButindex(sndIndex, sndData)
            Exit Sub
        
        Case SendTarget.ToClanArea
            Call SendToUserGuildArea(sndIndex, sndData)
            Exit Sub
        
        Case SendTarget.ToPartyArea
            Call SendToUserPartyArea(sndIndex, sndData)
            Exit Sub
        
        Case SendTarget.ToAdminsAreaButConsejeros
            Call SendToAdminsButConsejerosArea(sndIndex, sndData)
            Exit Sub
        
        Case SendTarget.ToNPCArea
            Call SendToNpcArea(sndIndex, sndData)
            Exit Sub
        
        Case SendTarget.ToDiosesYclan
            LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)

            While LoopC > 0

                If (UserList(LoopC).ConnID <> -1) Then
                    Call UserList(LoopC).outgoingData.WriteASCIIStringFixed(sndData)
                End If

                LoopC = modGuilds.m_Iterador_ProximoUserIndex(sndIndex)
            Wend
            
            LoopC = modGuilds.Iterador_ProximoGM(sndIndex)

            While LoopC > 0

                If (UserList(LoopC).ConnID <> -1) Then
                    Call UserList(LoopC).outgoingData.WriteASCIIStringFixed(sndData)
                End If

                LoopC = modGuilds.Iterador_ProximoGM(sndIndex)
            Wend
            
            Exit Sub
        
        Case SendTarget.ToConsejo

            For LoopC = 1 To LastUser

                If (UserList(LoopC).ConnID <> -1) Then
                    
                    If UserList(LoopC).flags.Privilegios And PlayerType.RoyalCouncil Then
                        Call UserList(LoopC).outgoingData.WriteASCIIStringFixed(sndData)
                    End If

                End If

            Next LoopC

            Exit Sub
        
        Case SendTarget.ToConsejoCaos

            For LoopC = 1 To LastUser

                If (UserList(LoopC).ConnID <> -1) Then
                    
                    If UserList(LoopC).flags.Privilegios And PlayerType.ChaosCouncil Then
                        Call UserList(LoopC).outgoingData.WriteASCIIStringFixed(sndData)
                    End If

                End If

            Next LoopC

            Exit Sub
        
        Case SendTarget.ToRolesMasters

            For LoopC = 1 To LastUser

                If (UserList(LoopC).ConnID <> -1) Then
                    
                    If UserList(LoopC).flags.Privilegios And PlayerType.RoleMaster Then
                        Call UserList(LoopC).outgoingData.WriteASCIIStringFixed(sndData)
                    End If

                End If

            Next LoopC

            Exit Sub
        
        Case SendTarget.ToCiudadanos

            For LoopC = 1 To LastUser

                If (UserList(LoopC).ConnID <> -1) Then
                    
                    If Not criminal(LoopC) Then
                        Call UserList(LoopC).outgoingData.WriteASCIIStringFixed(sndData)
                    End If

                End If

            Next LoopC

            Exit Sub
        
        Case SendTarget.ToCriminales

            For LoopC = 1 To LastUser

                If (UserList(LoopC).ConnID <> -1) Then
                    
                    If criminal(LoopC) Then
                        Call UserList(LoopC).outgoingData.WriteASCIIStringFixed(sndData)
                    End If

                End If

            Next LoopC

            Exit Sub
        
        Case SendTarget.ToReal

            For LoopC = 1 To LastUser

                If (UserList(LoopC).ConnID <> -1) Then
                    
                    If UserList(LoopC).Faccion.ArmadaReal = 1 Then
                        Call UserList(LoopC).outgoingData.WriteASCIIStringFixed(sndData)
                    End If

                End If

            Next LoopC

            Exit Sub
        
        Case SendTarget.ToCaos

            For LoopC = 1 To LastUser

                If (UserList(LoopC).ConnID <> -1) Then
                    
                    If UserList(LoopC).Faccion.FuerzasCaos = 1 Then
                        Call UserList(LoopC).outgoingData.WriteASCIIStringFixed(sndData)
                    End If

                End If

            Next LoopC

            Exit Sub
        
        Case SendTarget.ToCiudadanosYRMs

            For LoopC = 1 To LastUser

                If (UserList(LoopC).ConnID <> -1) Then
                    
                    If Not criminal(LoopC) Or (UserList(LoopC).flags.Privilegios And PlayerType.RoleMaster) <> 0 Then
                        Call UserList(LoopC).outgoingData.WriteASCIIStringFixed(sndData)
                    End If

                End If

            Next LoopC

            Exit Sub
        
        Case SendTarget.ToCriminalesYRMs

            For LoopC = 1 To LastUser

                If (UserList(LoopC).ConnID <> -1) Then
                    
                    If criminal(LoopC) Or (UserList(LoopC).flags.Privilegios And PlayerType.RoleMaster) <> 0 Then
                        Call UserList(LoopC).outgoingData.WriteASCIIStringFixed(sndData)
                    End If

                End If

            Next LoopC

            Exit Sub
        
        Case SendTarget.ToRealYRMs

            For LoopC = 1 To LastUser

                If (UserList(LoopC).ConnID <> -1) Then
                    
                    If UserList(LoopC).Faccion.ArmadaReal = 1 Or (UserList(LoopC).flags.Privilegios And PlayerType.RoleMaster) <> 0 Then
                        Call UserList(LoopC).outgoingData.WriteASCIIStringFixed(sndData)
                    End If

                End If

            Next LoopC

            Exit Sub
        
        Case SendTarget.ToCaosYRMs

            For LoopC = 1 To LastUser

                If (UserList(LoopC).ConnID <> -1) Then
                    
                    If UserList(LoopC).Faccion.FuerzasCaos = 1 Or (UserList(LoopC).flags.Privilegios And PlayerType.RoleMaster) <> 0 Then
                        Call UserList(LoopC).outgoingData.WriteASCIIStringFixed(sndData)
                    End If

                End If

            Next LoopC

            Exit Sub
        
        Case SendTarget.ToHigherAdmins

            For LoopC = 1 To LastUser

                If UserList(LoopC).ConnID <> -1 Then
                    
                    If UserList(LoopC).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then
                        Call UserList(LoopC).outgoingData.WriteASCIIStringFixed(sndData)
                    End If

                End If

            Next LoopC

            Exit Sub
            
        Case SendTarget.ToGMsAreaButRmsOrCounselors
            Call SendToGMsAreaButRmsOrCounselors(sndIndex, sndData)
            Exit Sub
            
        Case SendTarget.ToUsersAreaButGMs
            Call SendToUsersAreaButGMs(sndIndex, sndData)
            Exit Sub

        Case SendTarget.ToUsersAndRmsAndCounselorsAreaButGMs
            Call SendToUsersAndRmsAndCounselorsAreaButGMs(sndIndex, sndData)
            Exit Sub

    End Select

End Sub

Private Sub SendToUserArea(ByVal UserIndex As Integer, ByVal sdData As String)

    '**************************************************************
    'Author: Lucio N. Tourrilhes (DuNga)
    'Last Modify Date: Unknow
    '
    '**************************************************************

    Dim query() As Collision.UUID
    
    Dim i       As Long
    
    For i = 0 To ModAreas.QueryObservers(UserIndex, ENTITY_TYPE_PLAYER, query, ENTITY_TYPE_PLAYER)

        Debug.Print "SendToUserArea"
        Call UserList(query(i).Name).outgoingData.WriteASCIIStringFixed(sdData)

    Next i

End Sub

Private Sub SendToUserAreaButindex(ByVal UserIndex As Integer, ByVal sdData As String)

    '**************************************************************
    'Author: Lucio N. Tourrilhes (DuNga)
    'Last Modify Date: Unknow
    '
    '**************************************************************
    Dim query() As Collision.UUID

    Dim i       As Long
    
    For i = 0 To ModAreas.QueryObservers(UserIndex, ENTITY_TYPE_PLAYER, query, ENTITY_TYPE_PLAYER)

        If query(i).Name <> UserIndex Then
            Debug.Print "ToPCAreaButIndex"
            Call UserList(query(i).Name).outgoingData.WriteASCIIStringFixed(sdData)
        End If

    Next i

End Sub

Private Sub SendToDeadUserArea(ByVal UserIndex As Integer, ByVal sdData As String)

    '**************************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modify Date: Unknow
    '
    '**************************************************************
    
    Dim query() As Collision.UUID

    Dim i       As Long
    
    For i = 0 To ModAreas.QueryObservers(UserIndex, ENTITY_TYPE_PLAYER, query, ENTITY_TYPE_PLAYER)

        'Dead and admins read
        If UserList(query(i).Name).ConnIDValida = True And (UserList(query(i).Name).flags.Muerto = 1 Or (UserList(query(i).Name).flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero)) <> 0) Then
            Call UserList(query(i).Name).outgoingData.WriteASCIIStringFixed(sdData)
        End If

    Next i

End Sub

Private Sub SendToUserGuildArea(ByVal UserIndex As Integer, ByVal sdData As String)

    '**************************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modify Date: Unknow
    '
    '**************************************************************
    Dim query()    As Collision.UUID

    Dim i          As Long

    Dim GuildIndex As Integer
    
    GuildIndex = UserList(UserIndex).GuildIndex
    
    For i = 0 To ModAreas.QueryObservers(UserIndex, ENTITY_TYPE_PLAYER, query, ENTITY_TYPE_PLAYER)

        With UserList(query(i).Name)

            If (.GuildIndex = GuildIndex Or (.flags.Privilegios And PlayerType.Dios)) Then
                Call UserList(query(i).Name).outgoingData.WriteASCIIStringFixed(sdData)
            End If
            
        End With

    Next i

End Sub

Private Sub SendToUserPartyArea(ByVal UserIndex As Integer, ByVal sdData As String)

    '**************************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modify Date: Unknow
    '
    '**************************************************************
    Dim query()    As Collision.UUID

    Dim i          As Long

    Dim GroupIndex As Long

    GroupIndex = UserList(UserIndex).PartyIndex
    
    If GroupIndex = 0 Then Exit Sub
    
    For i = 0 To ModAreas.QueryObservers(UserIndex, ENTITY_TYPE_PLAYER, query, ENTITY_TYPE_PLAYER)

        If (UserList(query(i).Name).PartyIndex = GroupIndex) Then
            Call UserList(query(i).Name).outgoingData.WriteASCIIStringFixed(sdData)
        End If

    Next i

End Sub

Private Sub SendToAdminsButConsejerosArea(ByVal UserIndex As Integer, _
                                          ByVal sdData As String)

    '**************************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modify Date: Unknow
    '
    '**************************************************************
    Dim query()    As Collision.UUID

    Dim i          As Long

    Dim GroupIndex As Long

    GroupIndex = UserList(UserIndex).PartyIndex
    
    If GroupIndex = 0 Then Exit Sub
    
    For i = 0 To ModAreas.QueryObservers(UserIndex, ENTITY_TYPE_PLAYER, query, ENTITY_TYPE_PLAYER)
    
        If UserList(query(i).Name).flags.Privilegios And (PlayerType.SemiDios Or PlayerType.Dios Or PlayerType.Admin) Then
            Call UserList(query(i).Name).outgoingData.WriteASCIIStringFixed(sdData)
        End If

    Next i

End Sub

Private Sub SendToNpcArea(ByVal NPCIndex As Long, ByVal sdData As String)

    '**************************************************************
    'Author: Lucio N. Tourrilhes (DuNga)
    'Last Modify Date: Unknow
    '
    '**************************************************************
    Dim query() As Collision.UUID

    Dim i       As Long
    
    For i = 0 To ModAreas.QueryObservers(NPCIndex, ENTITY_TYPE_NPC, query, ENTITY_TYPE_PLAYER)
        Debug.Print "SendToNPCArea"
        Call UserList(query(i).Name).outgoingData.WriteASCIIStringFixed(sdData)
    Next i

End Sub

Public Sub SendToAreaByPos(ByVal Map As Integer, _
                           ByVal X As Integer, _
                           ByVal Y As Integer, _
                           ByVal sdData As String)

    '**************************************************************
    'Author: Lucio N. Tourrilhes (DuNga)
    'Last Modify Date: Unknow
    '
    '**************************************************************
    Dim query() As Collision.UUID

    Dim i       As Long

    Dim ItemID  As Long

    ItemID = Pack(Map, X, Y)
    
    For i = 0 To ModAreas.QueryObservers(ItemID, ENTITY_TYPE_OBJECT, query, ENTITY_TYPE_PLAYER)
    
        Call UserList(query(i).Name).outgoingData.WriteASCIIStringFixed(sdData)

    Next i

End Sub

Public Sub SendToMap(ByVal Map As Integer, ByVal sdData As String)

    '**************************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modify Date: 5/24/2007
    '
    '**************************************************************
    Dim LoopC     As Long

    Dim tempIndex As Integer

    For LoopC = 1 To LastUser
        
        If UserList(LoopC).Pos.Map = Map Then
        
            Call UserList(LoopC).outgoingData.WriteASCIIStringFixed(sdData)
        
        End If

    Next LoopC

End Sub

Private Sub SendToGMsAreaButRmsOrCounselors(ByVal UserIndex As Integer, _
                                            ByVal sdData As String)

    '**************************************************************
    'Author: Torres Patricio(Pato)
    'Last Modify Date: 12/02/2010
    '12/02/2010: ZaMa - Restrinjo solo a dioses, admins y gms.
    '15/02/2010: ZaMa - Cambio el nombre de la funcion (viejo: ToGmsArea, nuevo: ToGmsAreaButRMsOrCounselors)
    '**************************************************************
    Dim query() As Collision.UUID

    Dim i       As Long
    
    For i = 0 To ModAreas.QueryObservers(UserIndex, ENTITY_TYPE_PLAYER, query, ENTITY_TYPE_PLAYER)
        
        With UserList(query(i).Name)

            ' Exclusivo para dioses, admins y gms
            If (.flags.Privilegios And Not PlayerType.User And Not PlayerType.Consejero And Not PlayerType.RoleMaster) = .flags.Privilegios Then
                Call UserList(query(i).Name).outgoingData.WriteASCIIStringFixed(sdData)
            End If

        End With

    Next i

End Sub

Private Sub SendToUsersAreaButGMs(ByVal UserIndex As Integer, ByVal sdData As String)

    '**************************************************************
    'Author: Torres Patricio(Pato)
    'Last Modify Date: 10/17/2009
    '
    '**************************************************************
    
    Dim query() As Collision.UUID

    Dim i       As Long
    
    For i = 0 To ModAreas.QueryObservers(UserIndex, ENTITY_TYPE_PLAYER, query, ENTITY_TYPE_PLAYER)
        
        If UserList(query(i).Name).flags.Privilegios And PlayerType.User Then
            Call UserList(query(i).Name).outgoingData.WriteASCIIStringFixed(sdData)
        End If

    Next i

End Sub

Private Sub SendToUsersAndRmsAndCounselorsAreaButGMs(ByVal UserIndex As Integer, _
                                                     ByVal sdData As String)

    '**************************************************************
    'Author: Torres Patricio(Pato)
    'Last Modify Date: 10/17/2009
    '
    '**************************************************************
    Dim query() As Collision.UUID

    Dim i       As Long
    
    For i = 0 To ModAreas.QueryObservers(UserIndex, ENTITY_TYPE_PLAYER, query, ENTITY_TYPE_PLAYER)

        If UserList(query(i).Name).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster) Then
            Call UserList(query(i).Name).outgoingData.WriteASCIIStringFixed(sdData)
        End If

    Next i

End Sub

Public Sub AlertarFaccionarios(ByVal UserIndex As Integer)

    '**************************************************************
    'Author: ZaMa
    'Last Modify Date: 17/11/2009
    'Alerta a los faccionarios, dandoles una orientacion
    '**************************************************************
    
    Dim query()  As Collision.UUID
    Dim i        As Long
    Dim Font     As FontTypeNames
    Dim tempData As String
    
    If esCaos(UserIndex) Then
        Font = FontTypeNames.FONTTYPE_CONSEJOCAOS
    Else
        Font = FontTypeNames.FONTTYPE_CONSEJO
    End If
    
    For i = 0 To ModAreas.QueryObservers(UserIndex, ENTITY_TYPE_PLAYER, query, ENTITY_TYPE_PLAYER)
            
        If query(i).Name <> UserIndex Then

            ' Solo se envia a los de la misma faccion
            If SameFaccion(query(i).Name, query(i).Name) Then
                
                tempData = PrepareMessageConsoleMsg("Escuchas el llamado de un companero que proviene del " & GetDireccion(UserIndex, query(i).Name), Font)
                    
                Call UserList(query(i).Name).outgoingData.WriteASCIIStringFixed(tempData)

            End If

        End If

    Next i

End Sub

