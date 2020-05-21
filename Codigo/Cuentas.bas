Attribute VB_Name = "Cuentas"
Option Explicit

Public Sub LoginAccountDatabase(ByVal UserIndex As Integer, ByVal UserName As String)
    '***************************************************
    'Author: Lorwik
    'Last Modification: 20/05/2020
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query              As String

    Call Database_Connect

    query = "SELECT id, username, password, salt, hash, gemas, status FROM account "
    query = query & "WHERE UPPER(username) = '" & UCase$(UserName) & "';"

    Set Database_RecordSet = Database_Connection.Execute(query)

    If Database_RecordSet.BOF Or Database_RecordSet.EOF Then
        Call WriteErrorMsg(UserIndex, "Error al cargar la cuenta.")
        Call CloseUser(UserIndex)
        Exit Sub

    End If
    
    With UserList(UserIndex)
        
        'Guardo la información de la cuenta
        .AccountInfo.ID = CInt(Database_RecordSet!ID)
        .AccountInfo.UserName = Database_RecordSet!UserName
        .AccountInfo.password = Database_RecordSet!password
        .AccountInfo.salt = Database_RecordSet!salt
        .AccountInfo.Hash = Database_RecordSet!Hash
        .AccountInfo.Gemas = CLng(Database_RecordSet!Gemas)
        .AccountInfo.status = CBool(Database_RecordSet!status)
        
        Set Database_RecordSet = Nothing
    
        'Now the characters
        query = "SELECT id, name, level, gold, body_id, head_id, weapon_id, shield_id, helmet_id, race_id, class_id, pos_map, rep_average, is_dead FROM usuario "
        query = query & "WHERE account_id = " & .AccountInfo.ID & " AND deleted = FALSE;"
    
        Set Database_RecordSet = Database_Connection.Execute(query)
    
        .AccountInfo.NumChars = 0

        If Not Database_RecordSet.RecordCount = 0 Then
            Database_RecordSet.MoveFirst
    
            While Not Database_RecordSet.EOF
            
                'Incrementamos la cantidad de PJ creados actualmnente
                .AccountInfo.NumChars = .AccountInfo.NumChars + 1

                .AccountInfo.AccountPJ(.AccountInfo.NumChars).ID = Database_RecordSet!ID
                .AccountInfo.AccountPJ(.AccountInfo.NumChars).Name = Database_RecordSet!Name
                .AccountInfo.AccountPJ(.AccountInfo.NumChars).body = Database_RecordSet!body_id
                .AccountInfo.AccountPJ(.AccountInfo.NumChars).Head = Database_RecordSet!head_id
                .AccountInfo.AccountPJ(.AccountInfo.NumChars).weapon = Database_RecordSet!weapon_id
                .AccountInfo.AccountPJ(.AccountInfo.NumChars).shield = Database_RecordSet!shield_id
                .AccountInfo.AccountPJ(.AccountInfo.NumChars).helmet = Database_RecordSet!helmet_id
                .AccountInfo.AccountPJ(.AccountInfo.NumChars).Class = Database_RecordSet!class_id
                .AccountInfo.AccountPJ(.AccountInfo.NumChars).race = Database_RecordSet!race_id
                .AccountInfo.AccountPJ(.AccountInfo.NumChars).Map = Database_RecordSet!pos_map
                .AccountInfo.AccountPJ(.AccountInfo.NumChars).Level = Database_RecordSet!Level
                .AccountInfo.AccountPJ(.AccountInfo.NumChars).Gold = Database_RecordSet!Gold
                .AccountInfo.AccountPJ(.AccountInfo.NumChars).criminal = (Database_RecordSet!rep_average < 0)
                .AccountInfo.AccountPJ(.AccountInfo.NumChars).dead = Database_RecordSet!is_dead
                .AccountInfo.AccountPJ(.AccountInfo.NumChars).gameMaster = EsGmChar(Database_RecordSet!Name)
                Database_RecordSet.MoveNext
            Wend
    
        End If
    
        .flags.AccountLogged = True
        NumCuentas = NumCuentas + 1
        Call MostrarNumCuentas
    End With

    Set Database_RecordSet = Nothing
    Call Database_Close
    
    Call WriteUserAccountLogged(UserIndex)

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in LoginAccountDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub CloseAccount(ByVal UserIndex As Integer)
'*****************************************
'Autor: lorwik
'Fecha: 20/05/2020
'Descripcion: Borramos todos los datos almacenados de una cuenta
'*****************************************
    Dim i As Byte
    
    With UserList(UserIndex)
    
        .ConnIDValida = False
        .ConnID = -1
        
        'Guardo la información de la cuenta
        .AccountInfo.ID = 0
        .AccountInfo.UserName = vbNullString
        .AccountInfo.password = vbNullString
        .AccountInfo.salt = vbNullString
        .AccountInfo.Hash = vbNullString
        .AccountInfo.Gemas = 0
        .AccountInfo.status = False
        
        For i = 1 To .AccountInfo.NumChars
            .AccountInfo.AccountPJ(.AccountInfo.NumChars).ID = 0
            .AccountInfo.AccountPJ(.AccountInfo.NumChars).Name = vbNullString
            .AccountInfo.AccountPJ(.AccountInfo.NumChars).body = 0
            .AccountInfo.AccountPJ(.AccountInfo.NumChars).Head = 0
            .AccountInfo.AccountPJ(.AccountInfo.NumChars).weapon = 0
            .AccountInfo.AccountPJ(.AccountInfo.NumChars).shield = 0
            .AccountInfo.AccountPJ(.AccountInfo.NumChars).helmet = 0
            .AccountInfo.AccountPJ(.AccountInfo.NumChars).Class = 0
            .AccountInfo.AccountPJ(.AccountInfo.NumChars).race = 0
            .AccountInfo.AccountPJ(.AccountInfo.NumChars).Map = 0
            .AccountInfo.AccountPJ(.AccountInfo.NumChars).Level = 0
            .AccountInfo.AccountPJ(.AccountInfo.NumChars).Gold = 0
            .AccountInfo.AccountPJ(.AccountInfo.NumChars).criminal = False
            .AccountInfo.AccountPJ(.AccountInfo.NumChars).dead = False
            .AccountInfo.AccountPJ(.AccountInfo.NumChars).gameMaster = False
        Next i
        
        '¿Tiene algun personaje conectado?
        If .flags.UserLogged Then
            If NumUsers > 0 Then NumUsers = NumUsers - 1
            Call CloseUser(UserIndex)
        End If
        
        NumCuentas = NumCuentas - 1
        Call MostrarNumCuentas
        .flags.AccountLogged = False
        
    End With

End Sub

Public Function CuentaExisteDatabase(ByVal UserName As String) As Boolean

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 12/10/2018
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    Call Database_Connect

    query = "SELECT id FROM account WHERE UPPER(username) = '" & UCase$(UserName) & "';"

    Set Database_RecordSet = Database_Connection.Execute(query)

    If Database_RecordSet.BOF Or Database_RecordSet.EOF Then
        CuentaExisteDatabase = False
        Exit Function

    End If

    CuentaExisteDatabase = (Database_RecordSet.RecordCount > 0)
    Set Database_RecordSet = Nothing
    Call Database_Close

    Exit Function

ErrorHandler:
    Call LogDatabaseError("Error in CuentaExisteDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Function

Public Function CuentaVerificada(ByVal UserName As String) As Boolean

    '***************************************************
    'Author: Lorwik
    'Last Modification: 15/05/2020
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    Call Database_Connect

    query = "SELECT status FROM account WHERE UPPER(username) = '" & UCase$(UserName) & "';"

    Set Database_RecordSet = Database_Connection.Execute(query)

    If Database_RecordSet.BOF Or Database_RecordSet.EOF Then
       CuentaVerificada = False
        Exit Function

    End If

    CuentaVerificada = CBool(Database_RecordSet!status)

    Set Database_RecordSet = Nothing
    Call Database_Close

    Exit Function

ErrorHandler:
    Call LogDatabaseError("Error in CuentaVerificada: " & UserName & ". " & Err.Number & " - " & Err.description)

End Function

Public Function PersonajePerteneceCuenta(ByVal UserName As String, _
                                                 ByVal AccountHash As String) As Boolean

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 12/10/2018
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    Call Database_Connect

    query = "SELECT u.id FROM usuario u JOIN account a ON u.account_id = a.id WHERE UPPER(u.name) = '" & UCase$(UserName) & "' AND a.hash= '" & AccountHash & "';"

    Set Database_RecordSet = Database_Connection.Execute(query)

    If Database_RecordSet.BOF Or Database_RecordSet.EOF Then
        PersonajePerteneceCuenta = False
        Exit Function

    End If

    PersonajePerteneceCuenta = (Database_RecordSet.RecordCount > 0)
    Set Database_RecordSet = Nothing
    Call Database_Close

    Exit Function

ErrorHandler:
    Call LogDatabaseError("Error in PersonajePerteneceCuenta: " & UserName & ". " & Err.Number & " - " & Err.description)

End Function

Public Function GetCountUserAccount(ByVal HashAccount As String) As Byte

    '***************************************************
    'Author: Lorwik
    'Last Modification: 17/05/2020
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    Call Database_Connect

    query = "SELECT COUNT(*) FROM usuario WHERE deleted = 0 and account_id = (SELECT id FROM account WHERE hash = '" & HashAccount & "');"

    Set Database_RecordSet = Database_Connection.Execute(query)

    If Database_RecordSet.BOF Or Database_RecordSet.EOF Then
        GetCountUserAccount = 0
        Exit Function

    End If

    GetCountUserAccount = val(Database_RecordSet.Fields(0).Value)
    Set Database_RecordSet = Nothing
    Call Database_Close

    Exit Function
ErrorHandler:
    Call LogDatabaseError("Error in GetUserTrainingTimeDatabase: " & HashAccount & ". " & Err.Number & " - " & Err.description)

End Function

Public Function GetAccountSalt(ByVal AccountName As String) As String

    '***************************************************
    'Autor: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 20/09/2018
    'Get the user Password Salt
    '***************************************************

    GetAccountSalt = GetAccountSaltDatabase(AccountName)

End Function

Public Function GetUserSalt(ByVal UserName As String) As String

    '***************************************************
    'Autor: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 20/09/2018
    'Get the user Password Salt
    '***************************************************

    GetUserSalt = GetUserSaltDatabase(UserName)

End Function

Public Function GetAccountPassword(ByVal AccountName As String) As String

    '***************************************************
    'Autor: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 20/09/2018
    'Get the user Password
    '***************************************************

    GetAccountPassword = GetAccountPasswordDatabase(AccountName)

End Function

Public Function GetUserPassword(ByVal UserName As String) As String

    '***************************************************
    'Autor: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 20/09/2018
    'Get the user Password
    '***************************************************

    GetUserPassword = GetUserPasswordDatabase(UserName)

End Function

Public Function GetUserEmail(ByVal UserName As String) As String

    '***************************************************
    'Autor: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 20/09/2018
    'Get the user Email
    '***************************************************

    GetUserEmail = GetUserEmailDatabase(UserName)

End Function

Public Sub StorePasswordSalt(ByVal UserName As String, _
                             ByVal password As String, _
                             ByVal salt As String)

    '***************************************************
    'Autor: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 21/09/2018
    'Saves the password and salt
    '***************************************************

    Call StorePasswordSaltDatabase(UserName, password, salt)

End Sub

Public Sub SaveUserEmail(ByVal UserName As String, ByVal Email As String)

    '***************************************************
    'Autor: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 21/09/2018
    'Saves the email
    '***************************************************

    Call SaveUserEmailDatabase(UserName, Email)

End Sub

Public Sub SaveUserPunishment(ByVal UserName As String, _
                              ByVal Number As Integer, _
                              ByVal Reason As String)

    '***************************************************
    'Autor: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 21/09/2018
    'Saves a new punishment
    '***************************************************

    Call SaveUserPunishmentDatabase(UserName, Number, Reason)

End Sub

Public Sub AlterUserPunishment(ByVal UserName As String, _
                               ByVal Number As Integer, _
                               ByVal Reason As String)

    '***************************************************
    'Autor: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 21/09/2018
    'Saves a new punishment
    '***************************************************

    Call AlterUserPunishmentDatabase(UserName, Number, Reason)

End Sub

Public Sub ResetUserFacciones(ByVal UserName As String)

    '***************************************************
    'Autor: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 24/09/2018
    'Reset the imperial an legionary armies
    '***************************************************

    Call ResetUserFaccionesDatabase(UserName)

End Sub

Public Sub KickUserCouncils(ByVal UserName As String)

    '***************************************************
    'Autor: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 24/09/2018
    'Kicks the user from both councils
    '***************************************************

    Call KickUserCouncilsDatabase(UserName)

End Sub

Public Sub KickUserFacciones(ByVal UserName As String)

    '***************************************************
    'Autor: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 24/09/2018
    'Kicks the user from both factions
    '***************************************************

    Call KickUserFaccionesDatabase(UserName)

End Sub

Public Sub KickUserChaosLegion(ByVal UserName As String, ByVal KickerName As String)

    '***************************************************
    'Autor: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 24/09/2018
    'Kicks the user from ChaosLegion
    '***************************************************

    Call KickUserChaosLegionDatabase(UserName)

End Sub

Public Sub KickUserRoyalArmy(ByVal UserName As String, ByVal KickerName As String)

    '***************************************************
    'Autor: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 24/09/2018
    'Kicks the user from RoyalArmy
    '***************************************************

    Call KickUserRoyalArmyDatabase(UserName)

End Sub

Public Sub UpdateUserLogged(ByVal UserName As String, ByVal Logged As Byte)

    '***************************************************
    'Autor: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 24/09/2018
    'Updates the logged value for the user
    '***************************************************

    Call UpdateUserLoggedDatabase(UserName, Logged)

End Sub

Public Function GetUserLastIps(ByVal UserName As String) As String

    '***************************************************
    'Autor: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 20/09/2018
    'Get the user Last IPs list
    '***************************************************

    GetUserLastIps = GetUserLastIpsDatabase(UserName)

End Function

Public Function GetUserSkills(ByVal UserName As String) As String

    '***************************************************
    'Autor: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 20/09/2018
    'Get the user Skills list
    '***************************************************

    GetUserSkills = GetUserSkillsDatabase(UserName)

End Function

Public Function GetUserFreeSkills(ByVal UserName As String) As Integer

    '***************************************************
    'Autor: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 24/09/2018
    'Get the number of free skillspoints
    '***************************************************

    GetUserFreeSkills = GetUserFreeSkillsDatabase(UserName)

End Function

Public Sub SaveUserTrainingTime(ByVal UserName As String, ByVal trainingTime As Long)

    '***************************************************
    'Autor: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 24/09/2018
    'Updates the trainingTime value for the user
    '***************************************************

    Call SaveUserTrainingTimeDatabase(UserName, trainingTime)

End Sub

Public Function GetUserTrainingTime(ByVal UserName As String) As Long

    '***************************************************
    'Autor: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 24/09/2018
    'Get the training time in minutes
    '***************************************************

    GetUserTrainingTime = GetUserTrainingTimeDatabase(UserName)

End Function

Public Function UserBelongsToRoyalArmy(ByVal UserName As String) As Boolean

    '***************************************************
    'Autor: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 26/09/2018
    'Check if the user belongs to Royal Army
    '***************************************************

    UserBelongsToRoyalArmy = UserBelongsToRoyalArmyDatabase(UserName)

End Function

Public Function UserBelongsToChaosLegion(ByVal UserName As String) As Boolean

    '***************************************************
    'Autor: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 26/09/2018
    'Check if the user belongs to Chaos Legion
    '***************************************************

    UserBelongsToChaosLegion = UserBelongsToChaosLegionDatabase(UserName)

End Function

Public Function GetUserLevel(ByVal UserName As String) As Byte

    '***************************************************
    'Autor: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 26/09/2018
    'Get the User Level
    '***************************************************

    GetUserLevel = GetUserLevelDatabase(UserName)

End Function

Public Function GetUserPromedio(ByVal UserName As String) As Long

    '***************************************************
    'Autor: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 26/09/2018
    'Get the User Reputation Average
    '***************************************************

    GetUserPromedio = GetUserPromedioDatabase(UserName)

End Function

Public Function GetUserReenlists(ByVal UserName As String) As Byte

    '***************************************************
    'Autor: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 26/09/2018
    'Get the User Legion reenlists
    '***************************************************

    GetUserReenlists = GetUserReenlistsDatabase(UserName)

End Function

Public Sub SaveUserReenlists(ByVal UserName As String, ByVal Reenlists As Byte)

    '***************************************************
    'Autor: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 26/09/2018
    'Updates the number of reenlists
    '***************************************************

    Call SaveUserReenlistsDatabase(UserName, Reenlists)

End Sub

Public Sub SaveBan(ByVal UserName As String, _
                   ByVal Reason As String, _
                   ByVal BannedBy As String)

    '***************************************************
    'Autor: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 18/09/2018
    'Saves the ban flag and reason
    '***************************************************
    Call SaveBanDatabase(UserName, Reason, BannedBy)

End Sub

Public Function GetUserAmountOfPunishments(ByVal UserName As String) As Integer

    '***************************************************
    'Autor: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 19/09/2018
    'Get the user number of punishments
    '***************************************************
    
    GetUserAmountOfPunishments = GetUserAmountOfPunishmentsDatabase(UserName)

End Function

Public Sub SendUserPunishments(ByVal UserIndex As Integer, _
                               ByVal UserName As String, _
                               ByVal Count As Integer)

    '***************************************************
    'Autor: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 18/09/2018
    'Writes a console msg for each punishment
    '***************************************************

    Call SendUserPunishmentsDatabase(UserIndex, UserName, Count)

End Sub

Public Function GetUserPos(ByVal UserName As String) As String

    '***************************************************
    'Autor: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 19/09/2018
    'Get the user position
    '***************************************************
    
    GetUserPos = GetUserPosDatabase(UserName)

End Function
