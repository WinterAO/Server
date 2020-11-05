Attribute VB_Name = "Retos"
' Lautaro Leonel Marino. Lujan, Buenos Aires
' 13/08/2019
' Modulo de retos.
' Desde aca se manejan los retos desde 1vs1 hasta nVSn, configurable de forma facil.

Option Explicit

Public Const MAX_RETOS_SIMULTANEOS As Byte = 4

Public Enum eTipoReto
    None = 0
    FightOne = 1
    FightTwo = 2
    FightThree = 3
End Enum

Public Type tRetoUser
    UserIndex As Integer
    Team As Byte
    Rounds As Byte
End Type

Private Type tMapEvent
    Map As Integer
    X As Byte
    Y As Byte
    X2 As Byte
    Y2 As Byte
End Type

Private Type tRetos
    Run As Boolean
    Users() As tRetoUser
    RequiredGld As Long
End Type

Public Arenas(1 To MAX_RETOS_SIMULTANEOS) As tMapEvent
Public Retos(1 To MAX_RETOS_SIMULTANEOS) As tRetos

Public Sub LoadArenas()

    Dim i   As Long
    
    Dim RetosIO As clsIniManager
    Set RetosIO = New clsIniManager

    Call RetosIO.Initialize(DatPath & "Retos.dat")

    For i = LBound(Arenas) To UBound(Arenas)
        Arenas(i).Map = RetosIO.GetValue("ARENA" & CStr(i), "Mapa")
        Arenas(i).X = RetosIO.GetValue("ARENA" & CStr(i), "X")
        Arenas(i).X2 = RetosIO.GetValue("ARENA" & CStr(i), "X2")
        Arenas(i).Y = RetosIO.GetValue("ARENA" & CStr(i), "Y")
        Arenas(i).Y2 = RetosIO.GetValue("ARENA" & CStr(i), "Y2")
    Next
    
    Set RetosIO = Nothing
    
End Sub

Private Sub ResetDueloUser(ByVal UserIndex As Integer)

    On Error GoTo Error
    
        With UserList(UserIndex)

            If .Counters.TimeFight > 0 Then
                .Counters.TimeFight = 0
                Call WriteUserInEvent(UserIndex)
            End If
                          
            With Retos(.flags.SlotReto)
                .Users(UserList(UserIndex).flags.SlotRetoUser).UserIndex = 0
                .Users(UserList(UserIndex).flags.SlotRetoUser).Team = 0
                .Users(UserList(UserIndex).flags.SlotRetoUser).Rounds = 0
           End With
              
           .flags.SlotReto = 0
           .flags.SlotRetoUser = 255
           Call StatsDuelos(UserIndex)
           Call WarpPosAnt(UserIndex)

       End With
          
   Exit Sub

Error:

End Sub

Private Sub ResetDuelo(ByVal SlotReto As Byte)
    On Error GoTo Error

    Dim LoopC As Integer
          
        With Retos(SlotReto)
            For LoopC = LBound(.Users()) To UBound(.Users())
              
                If .Users(LoopC).UserIndex > 0 Then
                    ResetDueloUser .Users(LoopC).UserIndex
                End If
                  
                .Users(LoopC).UserIndex = 0
                .Users(LoopC).Rounds = 0
                .Users(LoopC).Team = 0

           Next LoopC
          
           .RequiredGld = 0
           .Run = False
       End With
          
   Exit Sub

Error:
       LogRetos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : ResetDuelo()"
End Sub

Private Function FreeSlotArena() As Byte
    Dim LoopC As Integer
          
    For LoopC = 1 To MAX_RETOS_SIMULTANEOS
        If Retos(LoopC).Run = False Then
            FreeSlotArena = LoopC
            Exit Function
        End If
    Next LoopC
    
End Function

Private Function FreeSlot() As Byte
    ' Slot libre para comenzar un nuevo enfrentamiento
    Dim LoopC As Integer
          
    FreeSlot = 0
          
    For LoopC = 1 To MAX_RETOS_SIMULTANEOS
        With Retos(LoopC)
            If .Run = False Then
                FreeSlot = LoopC
                Exit For
             End If
        End With
    Next LoopC
          
End Function

Private Sub PasateInteger(ByVal SlotArena As Byte, ByRef Users() As String)
    On Error GoTo Error

    ' Cuando se acepta un reto los UserId strings pasan a UserId integer
               
    With Retos(SlotArena)
        Dim LoopC As Integer
              
        ReDim .Users(LBound(Users()) To UBound(Users())) As tRetoUser
              
        For LoopC = LBound(.Users()) To UBound(.Users())
            .Users(LoopC).UserIndex = NameIndex(Users(LoopC))
                  
            If .Users(LoopC).UserIndex > 0 Then
                UserList(.Users(LoopC).UserIndex).Stats.Gld = UserList(.Users(LoopC).UserIndex).Stats.Gld - .RequiredGld
                Call WriteUpdateGold(.Users(LoopC).UserIndex)
            End If
                  
        Next LoopC
    End With
   Exit Sub

Error:
    LogRetos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : PasateInteger()"
End Sub

Private Sub RewardUsers(ByVal SlotReto As Byte, ByVal UserIndex As Integer)
    On Error GoTo Error
          
    Dim obj As obj
          
    With UserList(UserIndex)
        .Stats.Gld = .Stats.Gld + (Retos(SlotReto).RequiredGld * 2)
        Call WriteUpdateGold(UserIndex)
    End With
          
    Exit Sub

Error:
    LogRetos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : RewardUsers()"
End Sub

Private Function SetSubTipo(ByRef Users() As String) As eTipoReto
    On Error GoTo Error
          
    If UBound(Users()) = 1 Then
        SetSubTipo = FightOne
        Exit Function
    End If
          
    If UBound(Users()) = 3 Then
        SetSubTipo = FightTwo
        Exit Function
    End If
          
    If UBound(Users()) = 5 Then
        SetSubTipo = FightThree
        Exit Function
    End If
          
    SetSubTipo = 0
    Exit Function

Error:
    LogRetos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : SetSubTipo()"
End Function

Private Function CanSetUsers(ByRef Users() As String) As Boolean
    On Error GoTo Error
          
    Dim tUser As Integer
    Dim tmpUsers() As String
          
    Dim LoopC As Integer, loopX As Integer
    Dim Tmp As String
          
    ' Chequeos de cantidad de personajes teniendo en cuenta el tipo de reto.
        
    If SetSubTipo(Users()) = 0 Then
        CanSetUsers = False
        Exit Function
    End If
          
    ReDim tmpUsers(LBound(Users()) To UBound(Users())) As String
          
    For LoopC = LBound(Users()) To UBound(Users())
        tmpUsers(LoopC) = Users(LoopC)
    Next LoopC
          
          
    For LoopC = LBound(Users()) To UBound(Users())
        For loopX = LBound(Users()) To UBound(Users()) - LoopC
            If Not loopX = UBound(Users()) Then
                If StrComp(UCase$(tmpUsers(loopX)), UCase$(tmpUsers(loopX + 1))) = 0 Then
                    CanSetUsers = False
                    Exit Function
                Else
                    Tmp = tmpUsers(loopX)
                          
                    tmpUsers(loopX) = tmpUsers(loopX + 1)
                    tmpUsers(loopX + 1) = Tmp
                End If
            End If
        Next loopX
    Next LoopC
          
    CanSetUsers = True
    Exit Function

Error:
    LogRetos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : CanSetUsers()"
End Function

Private Function CanContinueFight(ByVal UserIndex As Integer) As Boolean
    On Error GoTo Error
          
    ' Si encontramos un personaje vivo el evento continua.
    Dim LoopC As Integer
    Dim SlotReto As Byte
    Dim SlotRetoUser As Byte
          
    SlotReto = UserList(UserIndex).flags.SlotReto
    SlotRetoUser = UserList(UserIndex).flags.SlotRetoUser

    CanContinueFight = False
          
    With Retos(SlotReto)
          
        For LoopC = LBound(.Users()) To UBound(.Users())
            If .Users(LoopC).UserIndex > 0 And .Users(LoopC).UserIndex <> UserIndex Then
                If .Users(SlotRetoUser).Team = .Users(LoopC).Team Then
                    With UserList(.Users(LoopC).UserIndex)
                        If .flags.Muerto = 0 Then
                            CanContinueFight = True
                            Exit Function
                        End If
                    End With
                End If
                              
            End If
        Next LoopC
              
    End With
    Exit Function

Error:
    LogRetos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : CanContinueFight()"
End Function

Private Function AttackerFight(ByVal SlotReto As Byte, ByVal TeamUser As Byte) As Integer
    On Error GoTo Error

    ' Buscamos al AttackerIndex (Caso abandono del evento)
    Dim LoopC As Integer
          
    With Retos(SlotReto)
        For LoopC = LBound(.Users()) To UBound(.Users())
            If .Users(LoopC).UserIndex > 0 Then
                If .Users(LoopC).Team > 0 And .Users(LoopC).Team <> TeamUser Then
                    AttackerFight = .Users(LoopC).UserIndex
                    Exit For
                End If
            End If
        Next LoopC
    End With
    Exit Function

Error:
    LogRetos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : AttackerFight()"
End Function

Private Function CanAcceptFight(ByVal UserIndex As Integer, _
                        ByVal UserName As String) As Boolean

    On Error GoTo Error
          
    ' Username es el que mando el reto al principio.
    ' Si esta online y cumple con los requisitos entra
    Dim SlotTemp As Byte
    Dim tUser As Integer
    Dim ArrayNulo As Long
          
    tUser = NameIndex(UserName)
              
    If tUser <= 0 Then
        ' Personaje offline
        Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
        CanAcceptFight = False
        Exit Function
    End If
              
    With UserList(tUser)
        'GetSafeArrayPointer .RetoTemp.Users, ArrayNulo
        'If ArrayNulo <= 0 Then Exit Function
                  
        SlotTemp = SearchFight(UCase$(UserList(UserIndex).Name), .RetoTemp.Users, .RetoTemp.Accepts)
                  
        If SlotTemp = 255 Then
            CanAcceptFight = False
            ' El personaje no te mando ninguna solicitud
            Exit Function
        End If
                  
        If .RetoTemp.Accepts(SlotTemp) = 1 Then
            ' El personaje ya acepta.
            CanAcceptFight = False
            Exit Function
        End If
                  
                  
        ' Valido el usuario
        .RetoTemp.Accepts(SlotTemp) = 1
        CanAcceptFight = True
                  
        ' ï¿½ Chequeo de aceptaciones
        If CheckAccepts(.RetoTemp.Accepts) Then
            GoFight tUser
        End If
          
          
    End With
              
    Exit Function

Error:
     LogRetos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : CanAcceptFight()"
End Function

Private Function ValidateFight_Users(ByVal UserIndex As Integer, _
                                    ByVal GldRequired As Long, _
                                    ByRef Users() As String) As Boolean
                                              
    On Error GoTo Error
          
    ' Validamos al Team seleccionado.
          
    Dim LoopC As Integer
    Dim tUser As Integer
                                     
    For LoopC = LBound(Users()) To UBound(Users())
        If Users(LoopC) <> vbNullString Then
            tUser = NameIndex(Users(LoopC))
                      
                      
            If tUser <= 0 Then
                'call SendMsjUsers("El personaje " & Users(LoopC) & " esta offline.", Users())
                Call WriteConsoleMsg(UserIndex, "El personaje " & Users(LoopC) & " esta offline", FontTypeNames.FONTTYPE_INFO)
                ValidateFight_Users = False
                Exit Function
            End If
            
            '¿Se invito a un GM?
            If EsGm(tUser) Then
                Call WriteConsoleMsg(UserIndex, "Los GMs no pueden participar en retos.", FontTypeNames.FONTTYPE_INFO)
                ValidateFight_Users = False
                Exit Function
            End If
                          
            With UserList(tUser)
                If .flags.Muerto = 1 Then
                    Call WriteConsoleMsg(UserIndex, "El personaje " & Users(LoopC) & " esta muerto.", FontTypeNames.FONTTYPE_INFO)
                    ValidateFight_Users = False
                    Exit Function
                End If
                              
                If MapInfo(.Pos.Map).Pk = True Then
                    ValidateFight_Users = False
                    Exit Function
                End If
                              
                If (.flags.SlotReto > 0) Then
                    Call WriteConsoleMsg(UserIndex, "El personaje " & Users(LoopC) & " esta participando en otro evento.", FontTypeNames.FONTTYPE_INFO)
                    ValidateFight_Users = False
                    Exit Function
                End If
                              
                If .flags.Comerciando Then
                    Call WriteConsoleMsg(UserIndex, "El personaje " & Users(LoopC) & " no esta disponible en este momento.", FontTypeNames.FONTTYPE_INFO)
                    ValidateFight_Users = False
                    Exit Function
                End If
                              
                If .Stats.Gld < GldRequired Then
                    Call WriteConsoleMsg(UserIndex, "El personaje " & .Name & " no tiene las monedas en su billetera.", FontTypeNames.FONTTYPE_INFO)
                    ValidateFight_Users = False
                    Exit Function
                End If
        
            End With
        End If
    Next LoopC
          
          
    ValidateFight_Users = True
          
    Exit Function

Error:
    LogRetos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : ValidateFight_Users()"
End Function

Private Function ValidateFight(ByVal UserIndex As Integer, _
                                ByVal GldRequired As Long, _
                                ByRef Users() As String) As Boolean
                                      
    On Error GoTo Error
          
    ' Validamos el enfrentamiento que se va a disputar
    ' UserIndex = Personaje que inicio la invitacion.
    '(Userindex, Tipo, GldRequired, Users) Then
              
    Dim LoopC As Integer
    Dim tUser As Integer

    If GldRequired < 0 Or GldRequired > 100000000 Then
        Call WriteConsoleMsg(UserIndex, "Oro Minimo: 0 . Oro Maximo 100.000.000", FontTypeNames.FONTTYPE_INFO)
        ValidateFight = False
        Exit Function
    End If
          
    ' Los Team estan diferentes en cuanto a cantidad. [LOG ERROR ANTI CHEAT]
    If Not CanSetUsers(Users) Then
        'Mensaje: Intento hackear el sistema
        Call LogRetos("POSIBLE HACKEO: " & UserList(UserIndex).Name & " hackeo el sistema de retos.")
        ValidateFight = False
        Exit Function
    End If
          
    ' Validamos a los personajes
    If Not ValidateFight_Users(UserIndex, GldRequired, Users()) Then
        ValidateFight = False
        Exit Function
    End If
          
          
    ValidateFight = True
          
    Exit Function

Error:
    LogRetos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : ValidateFight()"
End Function

Private Function StrTeam(ByRef Users() As tRetoUser) As String
          
    On Error GoTo Error
          
    ' Devuelve ENEMIGOS vs TEAM
          
    Dim LoopC As Integer
    Dim strtemp(1) As String
          
    ' 1 vs 1
    If UBound(Users()) = 1 Then
        If Users(0).UserIndex > 0 Then
            strtemp(0) = UserList(Users(0).UserIndex).Name
        Else
            strtemp(0) = "Usuario descalificado"
        End If
              
        If Users(1).UserIndex > 0 Then
            strtemp(1) = UserList(Users(1).UserIndex).Name
        Else
            strtemp(1) = "Usuario descalificado"
        End If
              
        StrTeam = strtemp(0) & " vs " & strtemp(1)
        Exit Function
    End If
          
    For LoopC = LBound(Users()) To UBound(Users())
        If Users(LoopC).UserIndex > 0 Then
            If LoopC < ((1 + UBound(Users)) / 2) Then
                strtemp(0) = strtemp(0) & UserList(Users(LoopC).UserIndex).Name & ", "
            Else
                strtemp(1) = strtemp(1) & UserList(Users(LoopC).UserIndex).Name & ", "
            End If
        End If
    Next LoopC
          
    If Not strtemp(0) = vbNullString Then
        strtemp(0) = Left$(strtemp(0), Len(strtemp(0)) - 2)
    Else
        strtemp(0) = "Equipo descalificado"
    End If
          
    If Not strtemp(1) = vbNullString Then
        strtemp(1) = Left$(strtemp(1), Len(strtemp(1)) - 2)
    Else
        strtemp(1) = "Equipo descalificado"
    End If
          
    StrTeam = strtemp(0) & " vs " & strtemp(1)
          
    Exit Function

Error:
    LogRetos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : StrTeam()"
End Function

Private Function CheckAccepts(ByRef Accepts() As Byte) As Boolean
    On Error GoTo Error
          
    ' Si encontramos a un usuario que no haya aceptado retornamos false.
    Dim LoopC As Integer
          
    CheckAccepts = True
          
    For LoopC = LBound(Accepts()) To UBound(Accepts())
        If Accepts(LoopC) = 0 Then
            CheckAccepts = False
            Exit Function
        End If
    Next LoopC
          
    Exit Function

Error:
    LogRetos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : CheckAccepts()"
End Function

Private Function SearchFight(ByVal UserName As String, _
                                ByRef Users() As String, _
                                ByRef Accepts() As Byte) As Byte
                                      
    ' Buscamos la invitacion que nos realizo el personaje UserName
          
    On Error GoTo Error

    Dim LoopC As Integer
          
    SearchFight = 255
          
    For LoopC = LBound(Users()) To UBound(Users())
        If StrComp(UCase$(Users(LoopC)), UCase$(UserName)) = 0 And Accepts(LoopC) = 0 Then
            SearchFight = LoopC
            Exit Function
        End If
    Next LoopC
          
    Exit Function

Error:
    LogRetos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : SearchFight()"
End Function
Public Function CanAttackReto(ByVal AttackerIndex As Integer, ByVal victimIndex As Integer) As Boolean
          
    On Error GoTo Error

    CanAttackReto = True
          
    With UserList(AttackerIndex)
        If .flags.SlotReto > 0 Then
                  
            'If Retos(.flags.SlotReto).Users(.flags.SlotRetoUser).Team = _
                Retos(.flags.SlotReto).Users(UserList(VictimIndex).flags.SlotRetoUser).Team Then
                    CanAttackReto = True
                    Exit Function
            'End If
        End If
          
    End With
          
    Exit Function

Error:
    LogRetos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : CanAttackReto()"
End Function

Private Sub SendInvitation(ByVal UserIndex As Integer, _
                            ByVal GldRequired As Long, _
                            ByRef Users() As String)
                                  
    On Error GoTo Error
          
    ' Enviamos la solicitud del duelo a los demas y guardamos los datos temporales al usuario mandatario.
          
    Dim LoopC As Integer
    Dim strtemp As String
    Dim tUser As Integer
    Dim Str() As tRetoUser
          
    ' Save data temp
    With UserList(UserIndex)
          
              
        With .RetoTemp
            ReDim .Accepts(LBound(Users()) To UBound(Users())) As Byte
            ReDim .Users(LBound(Users()) To UBound(Users())) As String
                  
            .RequiredGld = GldRequired
            .Users = Users
                  
            .Accepts(UBound(Users())) = 1 ' El ultimo personaje es el que enviï¿½ por lo tanto ya aceptï¿½.
        End With
    End With
          
    ReDim Str(LBound(Users()) To UBound(Users())) As tRetoUser
          
    For LoopC = LBound(Users()) To UBound(Users())
        Str(LoopC).UserIndex = NameIndex(Users(LoopC))
    Next LoopC
          
    strtemp = StrTeam(Str) & "."
    strtemp = strtemp & IIf(GldRequired > 0, " Oro requerido: " & GldRequired & ".", vbNullString)
    strtemp = strtemp & " Para aceptar tipea /ACEPTAR " & UserList(UserIndex).Name
          
    For LoopC = LBound(Users()) To UBound(Users())
        tUser = NameIndex(Users(LoopC))
              
        If tUser <> UserIndex Then
            Call WriteConsoleMsg(tUser, strtemp, FontTypeNames.FONTTYPE_WARNING)
        End If
                                              
    Next LoopC
          
    Exit Sub

Error:
    LogRetos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : SendInvitation()"
End Sub

Private Sub GoFight(ByVal UserIndex As Integer)
          ' Comienzo del duelo
          
    On Error GoTo Error

    Dim GldRequired As Long
    Dim SlotArena As Byte
          
    SlotArena = FreeSlotArena
          
    ' Mensaje : No hay mas arenas disponibles
    If SlotArena = 0 Then Exit Sub
          
    With UserList(UserIndex)
        If ValidateFight(UserIndex, .RetoTemp.RequiredGld, .RetoTemp.Users) Then
                  
            Retos(SlotArena).RequiredGld = .RetoTemp.RequiredGld
            Retos(SlotArena).Run = True
                  
            Call PasateInteger(SlotArena, .RetoTemp.Users)
                  
            Call SetUserEvent(SlotArena, Retos(SlotArena).Users)
            Call WarpFight(Retos(SlotArena).Users)
        End If
    End With
          
    Exit Sub

Error:
    LogRetos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : GoFight()"
End Sub

Private Sub SetUserEvent(ByVal SlotReto As Byte, ByRef Users() As tRetoUser)

    On Error GoTo Error
    ' Guardamos los slot en los usuarios y seteamos el team.
          
    Dim LoopC As Integer
    Dim SlotRetoUser As Byte
          
    For LoopC = LBound(Users()) To UBound(Users())
        If Users(LoopC).UserIndex > 0 Then
            With Users(LoopC)
                If .UserIndex > 0 Then
                    UserList(.UserIndex).flags.SlotReto = SlotReto
                    UserList(.UserIndex).flags.SlotRetoUser = LoopC
                          
                End If
            End With
              
            With Retos(SlotReto)
                If LoopC < ((1 + UBound(Users())) / 2) Then
                    .Users(LoopC).Team = 2
                Else
                    .Users(LoopC).Team = 1
                End If
            End With
              
            With UserList(Users(LoopC).UserIndex)
                .PosAnt.Map = .Pos.Map
                .PosAnt.X = .Pos.X
                .PosAnt.Y = .Pos.Y
                      
            End With
        End If
    Next LoopC
          
    Exit Sub

Error:
    LogRetos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : SetUserEvent()"
End Sub

Private Sub WarpFight(ByRef Users() As tRetoUser)

    ' Teletransportamos a los personajes a la sala de combate
          
    On Error GoTo Error

    Dim LoopC As Integer
    Dim tUser As Integer
    Dim Pos As WorldPos
    Const Tile_Extra As Byte = 5
          
    For LoopC = LBound(Users()) To UBound(Users())
        tUser = Users(LoopC).UserIndex
              
        If tUser > 0 Then
            Pos.Map = Arenas(UserList(tUser).flags.SlotReto).Map
                  
            If Users(LoopC).Team = 1 Then
                Pos.X = Arenas(UserList(tUser).flags.SlotReto).X
                Pos.Y = Arenas(UserList(tUser).flags.SlotReto).Y
            Else
                Pos.X = Arenas(UserList(tUser).flags.SlotReto).X2
                Pos.Y = Arenas(UserList(tUser).flags.SlotReto).Y2
            End If
                  
            With UserList(tUser)
                .Counters.TimeFight = 10
    
                Call WriteUserInEvent(tUser)
    
                ' Mensaje: Preparate en 10 segundos comenzarï¿½s a luchar!
                      
                Call ClosestStablePos(Pos, Pos)
                Call WarpUserChar(tUser, Pos.Map, Pos.X, Pos.Y, False)
            End With

        End If

    Next LoopC
          
    Exit Sub

Error:
    LogRetos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : WarpFight()"
End Sub

Private Sub AddRound(ByVal SlotReto As Byte, ByVal Team As Byte)

    On Error GoTo Error

    Dim LoopC As Integer
          
    With Retos(SlotReto)
        For LoopC = LBound(.Users()) To UBound(.Users())

            If .Users(LoopC).Team = Team And .Users(LoopC).UserIndex > 0 Then
                .Users(LoopC).Rounds = .Users(LoopC).Rounds + 1
            End If

        Next LoopC
          
    End With
          
    Exit Sub

Error:
    LogRetos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : AddRound()"
End Sub

Private Sub SendMsjUsers(ByVal strMsj As String, _
                        ByRef Users() As String)
                              
    On Error GoTo Error

    Dim LoopC As Integer
    Dim tUser As Integer
          
    For LoopC = LBound(Users()) To UBound(Users())

        tUser = NameIndex(Users(LoopC))

        If tUser > 0 Then
            Call WriteConsoleMsg(tUser, strMsj, FontTypeNames.FONTTYPE_VENENO)
        End If

    Next LoopC
          
    Exit Sub

Error:
    LogRetos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : SendMsjUsers()"
End Sub

Private Function ExistCompanero(ByVal UserIndex As Integer) As Boolean
          Dim LoopC As Integer
          Dim SlotReto As Byte
          Dim SlotRetoUser As Byte
          
   On Error GoTo ExistCompanero_Error

    SlotReto = UserList(UserIndex).flags.SlotReto
    SlotRetoUser = UserList(UserIndex).flags.SlotRetoUser
          
    With Retos(SlotReto)
        For LoopC = LBound(.Users()) To UBound(.Users())
            If .Users(LoopC).UserIndex > 0 Then
                If LoopC <> SlotRetoUser Then
                    If .Users(LoopC).Team = .Users(SlotRetoUser).Team Then
                        ExistCompanero = True
                        Exit For
                    End If
                End If
            End If
        Next LoopC
    End With

   On Error GoTo 0
   Exit Function

ExistCompanero_Error:

    LogRetos "Error " & Err.Number & " (" & Err.description & ") in procedure ExistCompanero of Modulo mRetos in line " & Erl
          
End Function

Public Sub UserDieFight(ByVal UserIndex As Integer, ByVal AttackerIndex As Integer, ByVal Forzado As Boolean)

    On Error GoTo Error

    ' Un personaje en reto es matado por otro.
    Dim LoopC As Integer
    Dim strtemp As String
    Dim SlotReto As Byte
    Dim TeamUser As Byte
    Dim Rounds As Byte
    Dim Deslogged As Boolean
    Dim ExistTeam As Boolean
          
    SlotReto = UserList(UserIndex).flags.SlotReto
          
    Deslogged = False
              
    ' Caso hipotetico de deslogeo. El funcionamiento es el mismo, con la diferencia de que se busca al ganador.
    If AttackerIndex = 0 Then
    AttackerIndex = AttackerFight(SlotReto, Retos(SlotReto).Users(UserList(UserIndex).flags.SlotRetoUser).Team)
    Deslogged = True
    End If
          
    TeamUser = Retos(SlotReto).Users(UserList(AttackerIndex).flags.SlotRetoUser).Team
    ExistTeam = ExistCompanero(UserIndex)
          
          
    ' Deslogeo de todos los integrantes del team
    If Forzado Then
        If Not ExistTeam Then
            Call FinishFight(SlotReto, TeamUser)
            Call ResetDuelo(SlotReto)
            Exit Sub
        End If
    End If
          
    With UserList(UserIndex)
        If Not CanContinueFight(UserIndex) Then

            With Retos(SlotReto)
    
                For LoopC = LBound(.Users()) To UBound(.Users())
    
                    With .Users(LoopC)
    
                        If .UserIndex > 0 And .Team = TeamUser Then
    
                            If Rounds = 0 Then
                                Call AddRound(SlotReto, .Team)
                                Rounds = .Rounds
                            End If
                                      
                            Call WriteConsoleMsg(.UserIndex, "Has ganado el round. Rounds ganados: " & .Rounds & ".", FontTypeNames.FONTTYPE_VENENO)
                                       
                        End If
    
                    End With
                              
                    If .Users(LoopC).UserIndex > 0 Then StatsDuelos .Users(LoopC).UserIndex
                Next LoopC
                          
                If Rounds >= (3 / 2) + 0.5 Or Forzado Then
                    Call FinishFight(SlotReto, TeamUser)
                    Call ResetDuelo(SlotReto)
                    Exit Sub
                Else
                    Call FinishFight(SlotReto, TeamUser, True)
                    'call StatsDuelos(Userindex)
                End If
    
            End With

        End If
              
 
    If Deslogged Then
        Call ResetDueloUser(UserIndex)
    End If

    End With
          
    Exit Sub

Error:
    LogRetos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : UserdieFight() en linea " & Erl
End Sub

Private Sub StatsDuelos(ByVal UserIndex As Integer)

    On Error GoTo Error

    With UserList(UserIndex)

        If .flags.Muerto Then
            Call RevivirUsuario(UserIndex)
            .Stats.MinHp = .Stats.MaxHp
            .Stats.MinMAN = .Stats.MaxMAN
            .Stats.MinSta = .Stats.MaxSta
                  
            Call WriteUpdateUserStats(UserIndex)
                    
            Exit Sub
        End If
                
        .Stats.MinHp = .Stats.MaxHp
        .Stats.MinMAN = .Stats.MaxMAN
        .Stats.MinSta = .Stats.MaxSta
                  
        Call WriteUpdateUserStats(UserIndex)
                
    End With
          
    Exit Sub

Error:
    LogRetos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : StatsDuelos()"
End Sub

Private Sub FinishFight(ByVal SlotReto As Byte, ByVal Team As Byte, Optional ByVal ChangeTeam As Boolean)

    ' Finalizamos el reto o el round.
          
    On Error GoTo Error

    Dim LoopC As Integer
    Dim strtemp As String
          
    With Retos(SlotReto)
    
        For LoopC = LBound(.Users()) To UBound(.Users())
            If .Users(LoopC).UserIndex > 0 Then
                If UserList(.Users(LoopC).UserIndex).Counters.TimeFight > 0 Then
                    UserList(.Users(LoopC).UserIndex).Counters.TimeFight = 0
                    WriteUserInEvent .Users(LoopC).UserIndex
                End If
                      
                If Team = .Users(LoopC).Team Then
                    If ChangeTeam Then
                        Call StatsDuelos(.Users(LoopC).UserIndex)
                    Else
                        .Run = False
                        Call StatsDuelos(.Users(LoopC).UserIndex)
                        Call RewardUsers(SlotReto, .Users(LoopC).UserIndex)
                                  
                        If .Users(LoopC).Rounds > 0 Then
                            Call WriteConsoleMsg(.Users(LoopC).UserIndex, "Has ganado el reto con " & .Users(LoopC).Rounds & " rounds a tu favor.", FontTypeNames.FONTTYPE_VENENO)
                        Else
                            Call WriteConsoleMsg(.Users(LoopC).UserIndex, "Has ganado el reto.", FontTypeNames.FONTTYPE_VENENO)
                        End If
    
                        strtemp = strtemp & UserList(.Users(LoopC).UserIndex).Name & ", "
                                  
                    End If
                          
                End If
            End If
        Next LoopC
          
        If ChangeTeam Then
            Call WarpFight(.Users())
        Else
            strtemp = Left$(strtemp, Len(strtemp) - 2)
            
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Retos: " & StrTeam(.Users()) & ". Ganador " & strtemp & ". Apuesta por " & .RequiredGld & " Monedas de Oro", FontTypeNames.FONTTYPE_INFO))
            Call LogRetos("Retos: " & StrTeam(.Users()) & ". Ganador el team de " & strtemp & ". Apuesta por " & .RequiredGld & " Monedas de Oro")
        End If
        
    End With
          
    Exit Sub

Error:
    LogRetos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : FinishFight() en linea " & Erl
End Sub

' Procedimientos necesarios para enviar, aceptar o abandonar.

Public Sub SendFight(ByVal UserIndex As Integer, _
                            ByVal GldRequired As Long, _
                            ByRef Users() As String)
          
    On Error GoTo Error
          
    ' Enviamos una solicitud de enfrentamiento
          
    With UserList(UserIndex)
              
        If ValidateFight(UserIndex, GldRequired, Users) Then
            Call SendInvitation(UserIndex, GldRequired, Users)
            Call WriteConsoleMsg(UserIndex, "Espera noticias para concretar el reto que has enviado. Recuerda que si vuelves a mandar, la anterior solicitud se cancela.", FontTypeNames.FONTTYPE_WARNING)
        End If
              
              
    End With
          
    Exit Sub
Error:
    LogRetos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : SendFight()"
End Sub

Public Sub AcceptFight(ByVal UserIndex As Integer, _
                        ByVal UserName As String)
                              
    On Error GoTo Error
                              
    With UserList(UserIndex)
              
        If CanAcceptFight(UserIndex, UserName) Then
                  
            Call WriteConsoleMsg(UserIndex, "Has aceptado la invitacion.", FontTypeNames.FONTTYPE_INFO)
            ' Has aceptado la invitacion bababa
        End If
        
    End With
          
    Exit Sub
Error:
    LogRetos "[" & Err.Number & "] " & Err.description & ") PROCEDIMIENTO : AcceptFight()"
End Sub

Public Sub WarpPosAnt(ByVal UserIndex As Integer)
    ' ï¿½ Warpeo del personaje a su posiciï¿½n anterior.
          
    Dim Pos As WorldPos
          
    On Error GoTo WarpPosAnt_Error

        With UserList(UserIndex)
            Pos.Map = .PosAnt.Map
            Pos.X = .PosAnt.X
            Pos.Y = .PosAnt.Y
                          
            Call FindLegalPos(UserIndex, Pos.Map, Pos.X, Pos.Y)
            Call WarpUserChar(UserIndex, Pos.Map, Pos.X, Pos.Y, False)
              
            .PosAnt.Map = 0
            .PosAnt.X = 0
            .PosAnt.Y = 0
          
        End With

   On Error GoTo 0
   Exit Sub

WarpPosAnt_Error:

    LogError "Error " & Err.Number & " (" & Err.description & ") in procedure WarpPosAnt of Modulo General in line " & Erl
End Sub

