Attribute VB_Name = "Cuentas"
Option Explicit

Public Sub LoginAccountDatabase(ByVal UserIndex As Integer, ByVal UserName As String)
    '***************************************************
    'Author: Lorwik
    'Last Modification: 20/05/2020
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query              As String
    Dim TieneGM            As Boolean
    
    Call Database_Connect

    query = "SELECT id, username, email, password, salt, gemas, status FROM account "
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
        .AccountInfo.Email = Database_RecordSet!Email
        .AccountInfo.Password = Database_RecordSet!Password
        .AccountInfo.salt = Database_RecordSet!salt
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
                
                If .AccountInfo.AccountPJ(.AccountInfo.NumChars).gameMaster = True Then TieneGM = True
                    
                Database_RecordSet.MoveNext
            Wend
    
        End If
        
        'Si el server esta restringido y no tiene GM no le dejamos entrar.
        If ServerSoloGMs <> 0 And TieneGM = False Then
            Call WriteErrorMsg(UserIndex, "El servidor se encuentra en estos momentos en mantenimiento. Intentelo mas tarde.")
            Call CloseUser(UserIndex)
            Exit Sub
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
        .AccountInfo.Password = vbNullString
        .AccountInfo.salt = vbNullString
        .AccountInfo.Gemas = 0
        .AccountInfo.status = False
        
        'Reseteo la IP
        .IP = vbNullString
        
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
            Call Cerrar_Usuario(UserIndex)
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

Public Function PersonajePerteneceCuenta(ByVal UserIndex As Integer, ByVal UserName As String) As Boolean

    '***************************************************
    'Author: Lorwik
    'Last Modification: 04/06/2020
    'Descripcion: Comprobamos si el personaje pertenece a la cuenta, para ello
    'hacemos una consulta buscando el nombre del personaje y el account_id de
    'la persona que quieres entrar al personajeque le pasamos, si obtenemos 1 resultado
    'el personaje pertenece a la cuenta
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    Call Database_Connect

    query = "SELECT id FROM usuario WHERE UPPER(name) = '" & UCase$(UserName) & "' AND account_id = '" & UserList(UserIndex).AccountInfo.ID & "';"

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

Public Function GetCountUserAccount(ByVal UserIndex As Integer) As Byte

    '***************************************************
    'Author: Lorwik
    'Last Modification: 04/06/2020
    'Descripcion: Comprobamos la cantidad de personajes creados en la cuenta
    'para ello hacemos una consulta en la que buscamos todos los personajes
    'asociados al id de cuenta del UserIndex
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    Call Database_Connect

    query = "SELECT COUNT(*) FROM usuario WHERE deleted = 0 and account_id = '" & UserList(UserIndex).AccountInfo.ID & "';"

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
    Call LogDatabaseError("Error in GetUserTrainingTimeDatabase: UserIndex: " & UserIndex & " - Hash: " & UserList(UserIndex).AccountInfo.ID & ". " & Err.Number & " - " & Err.description)

End Function

Public Sub BorrarUsuarioDatabase(ByVal UserName As String)

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 10/10/2018
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    Call Database_Connect

    query = "UPDATE usuario SET name = '" & UCase$(UserName) & "_deleted', deleted = TRUE WHERE UPPER(name) = '" & UCase$(UserName) & "';"

    Database_Connection.Execute (query)

    Call Database_Close

    Exit Sub

ErrorHandler:
    Call LogDatabaseError("Error in BorrarUsuarioDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Function GetAccountSalt(ByVal AccountName As String) As String

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 10/10/2018
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    Call Database_Connect

    query = "SELECT salt FROM account WHERE UPPER(username) = '" & UCase$(AccountName) & "';"

    Set Database_RecordSet = Database_Connection.Execute(query)

    If Database_RecordSet.BOF Or Database_RecordSet.EOF Then
        GetAccountSalt = vbNullString
        Exit Function

    End If

    GetAccountSalt = Database_RecordSet!salt
    Set Database_RecordSet = Nothing
    Call Database_Close

    Exit Function
ErrorHandler:
    Call LogDatabaseError("Error in GetAccountSalt: " & AccountName & ". " & Err.Number & " - " & Err.description)

End Function

Public Function GetUserSalt(ByVal UserName As String) As String

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 10/10/2018
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    Call Database_Connect

    query = "SELECT salt FROM account WHERE id = (SELECT account_id FROM usuario WHERE UPPER(name) = '" & UCase$(UserName) & "');"

    Set Database_RecordSet = Database_Connection.Execute(query)

    If Database_RecordSet.BOF Or Database_RecordSet.EOF Then
        GetUserSalt = vbNullString
        Exit Function

    End If

    GetUserSalt = Database_RecordSet!salt
    Set Database_RecordSet = Nothing
    Call Database_Close

    Exit Function
ErrorHandler:
    Call LogDatabaseError("Error in GetUserSalt: " & UserName & ". " & Err.Number & " - " & Err.description)

End Function

Public Function GetAccountPassword(ByVal AccountName As String) As String

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 10/10/2018
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    Call Database_Connect

    query = "SELECT password FROM account WHERE UPPER(username) = '" & UCase$(AccountName) & "';"

    Set Database_RecordSet = Database_Connection.Execute(query)

    If Database_RecordSet.BOF Or Database_RecordSet.EOF Then
        GetAccountPassword = vbNullString
        Exit Function

    End If

    GetAccountPassword = Database_RecordSet!Password
    Set Database_RecordSet = Nothing
    Call Database_Close

    Exit Function
ErrorHandler:
    Call LogDatabaseError("Error in GetAccountPassword: " & AccountName & ". " & Err.Number & " - " & Err.description)

End Function

Public Function GetUserPassword(ByVal UserName As String) As String

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 10/10/2018
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    Call Database_Connect

    query = "SELECT password FROM account WHERE id = (SELECT account_id FROM usuario WHERE UPPER(name) = '" & UCase$(UserName) & "');"

    Set Database_RecordSet = Database_Connection.Execute(query)

    If Database_RecordSet.BOF Or Database_RecordSet.EOF Then
        GetUserPassword = vbNullString
        Exit Function

    End If

    GetUserPassword = Database_RecordSet!Password
    Set Database_RecordSet = Nothing
    Call Database_Close

    Exit Function
ErrorHandler:
    Call LogDatabaseError("Error in GetUserPassword: " & UserName & ". " & Err.Number & " - " & Err.description)

End Function

Public Function GetUserEmail(ByVal UserName As String) As String

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 10/10/2018
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    Call Database_Connect

    query = "SELECT username FROM account WHERE id = (SELECT account_id FROM usuario WHERE UPPER(name) = '" & UCase$(UserName) & "');"

    Set Database_RecordSet = Database_Connection.Execute(query)

    If Database_RecordSet.BOF Or Database_RecordSet.EOF Then
        GetUserEmail = vbNullString
        Exit Function

    End If

    GetUserEmail = Database_RecordSet!UserName
    Set Database_RecordSet = Nothing
    Call Database_Close

    Exit Function
ErrorHandler:
    Call LogDatabaseError("Error in GetUserEmail: " & UserName & ". " & Err.Number & " - " & Err.description)

End Function
