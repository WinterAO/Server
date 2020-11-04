Attribute VB_Name = "Cuentas"
Option Explicit

Public Sub LoginAccountDatabase(ByVal UserIndex As Integer, ByVal UserName As String, Optional ByVal Refresh As Boolean = False)
    '***************************************************
    'Author: Lorwik
    'Last Modification: 20/05/2020
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query              As String
    Dim TieneGM            As Boolean
    
    #If DBConexionUnica = 0 Then
        Call Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If CheckSQLStatus = False Then Database_Reconnect
    #End If
    
    With UserList(UserIndex)
    
    If Refresh = False Then
    
        query = "SELECT id, username, email, password, salt, gemas, status FROM account "
        query = query & "WHERE UPPER(username) = '" & UCase$(UserName) & "';"
    
        Set Database_RecordSet = Database_Connection.Execute(query)
    
        If Database_RecordSet.BOF Or Database_RecordSet.EOF Then
            Call WriteErrorMsg(UserIndex, "Error al cargar la cuenta.")
            Call CloseUser(UserIndex)
            Exit Sub
    
        End If
        
        'Guardo la información de la cuenta
        .AccountInfo.ID = CInt(Database_RecordSet!ID)
        .AccountInfo.UserName = Database_RecordSet!UserName
        .AccountInfo.Email = Database_RecordSet!Email
        .AccountInfo.Password = Database_RecordSet!Password
        .AccountInfo.Salt = Database_RecordSet!Salt
        .AccountInfo.Gemas = CLng(Database_RecordSet!Gemas)
        .AccountInfo.status = CBool(Database_RecordSet!status)
        
        Set Database_RecordSet = Nothing
        
    End If
    
        'Now the characters
        query = "SELECT id, name, level, body_id, head_id, weapon_id, shield_id, helmet_id, race_id, class_id, pos_map, rep_average, is_dead FROM usuario "
        query = query & "WHERE account_id = " & .AccountInfo.ID & " AND deleted = FALSE;"
    
        Set Database_RecordSet = Database_Connection.Execute(query)
    
        .AccountInfo.NumPjs = 0

        If Not Database_RecordSet.RecordCount = 0 Then
            Database_RecordSet.MoveFirst
    
            While Not Database_RecordSet.EOF
            
                'Incrementamos la cantidad de PJ creados actualmnente
                .AccountInfo.NumPjs = .AccountInfo.NumPjs + 1

                .AccountInfo.AccountPJ(.AccountInfo.NumPjs).ID = Database_RecordSet!ID
                .AccountInfo.AccountPJ(.AccountInfo.NumPjs).Name = Database_RecordSet!Name
                .AccountInfo.AccountPJ(.AccountInfo.NumPjs).body = Database_RecordSet!body_id
                .AccountInfo.AccountPJ(.AccountInfo.NumPjs).Head = Database_RecordSet!head_id
                .AccountInfo.AccountPJ(.AccountInfo.NumPjs).weapon = Database_RecordSet!weapon_id
                .AccountInfo.AccountPJ(.AccountInfo.NumPjs).shield = Database_RecordSet!shield_id
                .AccountInfo.AccountPJ(.AccountInfo.NumPjs).helmet = Database_RecordSet!helmet_id
                .AccountInfo.AccountPJ(.AccountInfo.NumPjs).Class = Database_RecordSet!class_id
                .AccountInfo.AccountPJ(.AccountInfo.NumPjs).race = Database_RecordSet!race_id
                .AccountInfo.AccountPJ(.AccountInfo.NumPjs).Map = Database_RecordSet!pos_map
                .AccountInfo.AccountPJ(.AccountInfo.NumPjs).level = Database_RecordSet!level
                .AccountInfo.AccountPJ(.AccountInfo.NumPjs).criminal = (Database_RecordSet!rep_average < 0)
                .AccountInfo.AccountPJ(.AccountInfo.NumPjs).dead = Database_RecordSet!is_dead
                .AccountInfo.AccountPJ(.AccountInfo.NumPjs).gameMaster = EsGmChar(Database_RecordSet!Name)
                
                If .AccountInfo.AccountPJ(.AccountInfo.NumPjs).gameMaster = True Then TieneGM = True
                    
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
    
    #If DBConexionUnica = 0 Then
        Call Database_Close
    #End If
    
    Call WriteEnviarPJUserAccount(UserIndex, Refresh)

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in LoginAccountDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub CloseAccount(ByVal UserIndex As Integer)
'*****************************************
'Autor: lorwik
'Fecha: 20/05/2020
'Descripcion: Cierra la cuenta
'*****************************************

    With UserList(UserIndex)
    
        .ConnIDValida = False
        .ConnID = -1
        
        Call ResetUseRaccount(UserIndex)
        
        '¿Tiene algun personaje conectado?
        If .flags.UserLogged Then
            Call Cerrar_Usuario(UserIndex)
        End If
        
        'Reseteo la IP
        .IP = vbNullString
        
        If NumCuentas > 0 Then NumCuentas = NumCuentas - 1
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

    #If DBConexionUnica = 0 Then
        Call Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If CheckSQLStatus = False Then Database_Reconnect
    #End If

    query = "SELECT id FROM account WHERE UPPER(username) = '" & UCase$(UserName) & "';"

    Set Database_RecordSet = Database_Connection.Execute(query)

    If Database_RecordSet.BOF Or Database_RecordSet.EOF Then
        CuentaExisteDatabase = False
        Exit Function

    End If

    CuentaExisteDatabase = (Database_RecordSet.RecordCount > 0)
    Set Database_RecordSet = Nothing
    
    #If DBConexionUnica = 0 Then
        Call Database_Close
    #End If

    Exit Function

ErrorHandler:
    If Err.Number = -1207576359 Then _
        Call Database_Reconnect

    Call LogDatabaseError("Error in CuentaExisteDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Function

Public Function CuentaVerificada(ByVal UserName As String) As Boolean

    '***************************************************
    'Author: Lorwik
    'Last Modification: 15/05/2020
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    #If DBConexionUnica = 0 Then
        Call Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If CheckSQLStatus = False Then Database_Reconnect
    #End If

    query = "SELECT status FROM account WHERE UPPER(username) = '" & UCase$(UserName) & "';"

    Set Database_RecordSet = Database_Connection.Execute(query)

    If Database_RecordSet.BOF Or Database_RecordSet.EOF Then
       CuentaVerificada = False
        Exit Function

    End If

    CuentaVerificada = CBool(Database_RecordSet!status)

    Set Database_RecordSet = Nothing
    
    #If DBConexionUnica = 0 Then
        Call Database_Close
    #End If

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

    #If DBConexionUnica = 0 Then
        Call Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If CheckSQLStatus = False Then Database_Reconnect
    #End If

    query = "SELECT id FROM usuario WHERE UPPER(name) = '" & UCase$(UserName) & "' AND account_id = '" & UserList(UserIndex).AccountInfo.ID & "';"

    Set Database_RecordSet = Database_Connection.Execute(query)

    If Database_RecordSet.BOF Or Database_RecordSet.EOF Then
        PersonajePerteneceCuenta = False
        Exit Function

    End If

    PersonajePerteneceCuenta = (Database_RecordSet.RecordCount > 0)
    Set Database_RecordSet = Nothing
    
    #If DBConexionUnica = 0 Then
        Call Database_Close
    #End If

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

    #If DBConexionUnica = 0 Then
        Call Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If CheckSQLStatus = False Then Database_Reconnect
    #End If

    query = "SELECT COUNT(*) FROM usuario WHERE deleted = 0 and account_id = '" & UserList(UserIndex).AccountInfo.ID & "';"

    Set Database_RecordSet = Database_Connection.Execute(query)

    If Database_RecordSet.BOF Or Database_RecordSet.EOF Then
        GetCountUserAccount = 0
        Exit Function

    End If

    GetCountUserAccount = val(Database_RecordSet.Fields(0).Value)
    Set Database_RecordSet = Nothing
    
    #If DBConexionUnica = 0 Then
        Call Database_Close
    #End If

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

    #If DBConexionUnica = 0 Then
        Call Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If CheckSQLStatus = False Then Database_Reconnect
    #End If

    query = "UPDATE usuario SET name = '" & UCase$(UserName) & "_deleted', deleted = TRUE WHERE UPPER(name) = '" & UCase$(UserName) & "';"

    Database_Connection.Execute (query)

    #If DBConexionUnica = 0 Then
        Call Database_Close
    #End If

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

    #If DBConexionUnica = 0 Then
        Call Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If CheckSQLStatus = False Then Database_Reconnect
    #End If

    query = "SELECT salt FROM account WHERE UPPER(username) = '" & UCase$(AccountName) & "';"

    Set Database_RecordSet = Database_Connection.Execute(query)

    If Database_RecordSet.BOF Or Database_RecordSet.EOF Then
        GetAccountSalt = vbNullString
        Exit Function

    End If

    GetAccountSalt = Database_RecordSet!Salt
    Set Database_RecordSet = Nothing
    
    #If DBConexionUnica = 0 Then
        Call Database_Close
    #End If

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

    #If DBConexionUnica = 0 Then
        Call Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If CheckSQLStatus = False Then Database_Reconnect
    #End If

    query = "SELECT salt FROM account WHERE id = (SELECT account_id FROM usuario WHERE UPPER(name) = '" & UCase$(UserName) & "');"

    Set Database_RecordSet = Database_Connection.Execute(query)

    If Database_RecordSet.BOF Or Database_RecordSet.EOF Then
        GetUserSalt = vbNullString
        Exit Function

    End If

    GetUserSalt = Database_RecordSet!Salt
    Set Database_RecordSet = Nothing
    
    #If DBConexionUnica = 0 Then
        Call Database_Close
    #End If

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

    #If DBConexionUnica = 0 Then
        Call Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If CheckSQLStatus = False Then Database_Reconnect
    #End If

    query = "SELECT password FROM account WHERE UPPER(username) = '" & UCase$(AccountName) & "';"

    Set Database_RecordSet = Database_Connection.Execute(query)

    If Database_RecordSet.BOF Or Database_RecordSet.EOF Then
        GetAccountPassword = vbNullString
        Exit Function

    End If

    GetAccountPassword = Database_RecordSet!Password
    Set Database_RecordSet = Nothing
    
    #If DBConexionUnica = 0 Then
        Call Database_Close
    #End If

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

    #If DBConexionUnica = 0 Then
        Call Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If CheckSQLStatus = False Then Database_Reconnect
    #End If

    query = "SELECT password FROM account WHERE id = (SELECT account_id FROM usuario WHERE UPPER(name) = '" & UCase$(UserName) & "');"

    Set Database_RecordSet = Database_Connection.Execute(query)

    If Database_RecordSet.BOF Or Database_RecordSet.EOF Then
        GetUserPassword = vbNullString
        Exit Function

    End If

    GetUserPassword = Database_RecordSet!Password
    Set Database_RecordSet = Nothing
    
    #If DBConexionUnica = 0 Then
        Call Database_Close
    #End If

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

    #If DBConexionUnica = 0 Then
        Call Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If CheckSQLStatus = False Then Database_Reconnect
    #End If

    query = "SELECT username FROM account WHERE id = (SELECT account_id FROM usuario WHERE UPPER(name) = '" & UCase$(UserName) & "');"

    Set Database_RecordSet = Database_Connection.Execute(query)

    If Database_RecordSet.BOF Or Database_RecordSet.EOF Then
        GetUserEmail = vbNullString
        Exit Function

    End If

    GetUserEmail = Database_RecordSet!UserName
    Set Database_RecordSet = Nothing
    
    #If DBConexionUnica = 0 Then
        Call Database_Close
    #End If

    Exit Function
ErrorHandler:
    Call LogDatabaseError("Error in GetUserEmail: " & UserName & ". " & Err.Number & " - " & Err.description)

End Function

Public Function SaveNewAccount(ByVal UserName As String, _
                                  ByVal Email As String, _
                                  ByVal Password As String, _
                                  ByVal Salt As String) As Boolean

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 12/10/2018
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    'Si perdimos la conexion reconectamos
    If CheckSQLStatus = False Then Database_Reconnect

    query = "INSERT INTO account SET "
    query = query & "username = '" & UserName & "', "
    query = query & "email = '" & Email & "', "
    query = query & "password = '" & Password & "', "
    query = query & "salt = '" & Salt & "', "
    query = query & "id_confirmacion = 'VERIFICADA', "
    query = query & "status = '1', "
    query = query & "date_created = NOW(), "
    query = query & "date_last_login = NOW();"

    Database_Connection.Execute (query)

    SaveNewAccount = True
    
    Exit Function
ErrorHandler:
    Call LogDatabaseError("Error in SaveNewAccountDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)
    SaveNewAccount = False

End Function

Public Sub ActualizarPJCuentas(ByVal UserIndex As Integer)
'****************************************************
'Autor: Lorwik
'Fecha: 04/11/2020
'Descripcion: Actualiza el PJ actual en el listado de la cuenta y lo manda al cliente
'****************************************************

    Dim Posicion As Byte
    Dim i As Byte

    With UserList(UserIndex)
    
        For i = 1 To .AccountInfo.NumPjs
            If .AccountInfo.AccountPJ(i).ID = .ID Then _
                Posicion = i
        Next i

        .AccountInfo.AccountPJ(Posicion).ID = .ID
        .AccountInfo.AccountPJ(Posicion).Name = .Name
        .AccountInfo.AccountPJ(Posicion).body = .Char.body
        .AccountInfo.AccountPJ(Posicion).Head = .Char.Head
        .AccountInfo.AccountPJ(Posicion).weapon = .Char.WeaponAnim
        .AccountInfo.AccountPJ(Posicion).shield = .Char.ShieldAnim
        .AccountInfo.AccountPJ(Posicion).helmet = .Char.CascoAnim
        .AccountInfo.AccountPJ(Posicion).Class = .clase
        .AccountInfo.AccountPJ(Posicion).race = .Raza
        .AccountInfo.AccountPJ(Posicion).Map = .Pos.Map
        .AccountInfo.AccountPJ(Posicion).level = .Stats.ELV
        .AccountInfo.AccountPJ(Posicion).criminal = criminal(UserIndex)
        .AccountInfo.AccountPJ(Posicion).dead = .flags.Muerto
        .AccountInfo.AccountPJ(Posicion).gameMaster = EsGmChar(.Name)
            
        'Actualiza los PJ de la cuenta
        Call WriteEnviarPJUserAccount(UserIndex, True)
    End With

End Sub

Public Sub AddNewPJCuenta(ByVal UserIndex As Integer)
'****************************************************
'Autor: Lorwik
'Fecha: 04/11/2020
'Descripcion: Añade un nuevo personaje a la lista de la cuenta
'****************************************************

    With UserList(UserIndex)
        'Incrementamos la cantidad de PJ creados actualmnente
        .AccountInfo.NumPjs = .AccountInfo.NumPjs + 1
            
        .AccountInfo.AccountPJ(.AccountInfo.NumPjs).ID = .ID
        .AccountInfo.AccountPJ(.AccountInfo.NumPjs).Name = .Name
        .AccountInfo.AccountPJ(.AccountInfo.NumPjs).body = .Char.body
        .AccountInfo.AccountPJ(.AccountInfo.NumPjs).Head = .Char.Head
        .AccountInfo.AccountPJ(.AccountInfo.NumPjs).weapon = .Char.WeaponAnim
        .AccountInfo.AccountPJ(.AccountInfo.NumPjs).shield = .Char.ShieldAnim
        .AccountInfo.AccountPJ(.AccountInfo.NumPjs).helmet = .Char.CascoAnim
        .AccountInfo.AccountPJ(.AccountInfo.NumPjs).Class = .clase
        .AccountInfo.AccountPJ(.AccountInfo.NumPjs).race = .Raza
        .AccountInfo.AccountPJ(.AccountInfo.NumPjs).Map = .Pos.Map
        .AccountInfo.AccountPJ(.AccountInfo.NumPjs).level = .Stats.ELV
        .AccountInfo.AccountPJ(.AccountInfo.NumPjs).criminal = criminal(UserIndex)
        .AccountInfo.AccountPJ(.AccountInfo.NumPjs).dead = .flags.Muerto
        .AccountInfo.AccountPJ(.AccountInfo.NumPjs).gameMaster = EsGmChar(.Name)
        
    End With
    
End Sub
