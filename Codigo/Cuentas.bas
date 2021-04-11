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
    
    #If DBConexionUnica = 0 Then
        Call Account_Database.Database_Connect
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If Account_Database.CheckSQLStatus = False Then Account_Database.Database_Reconnect
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If
    
    With UserList(UserIndex)
        
        '***********************
        'LOGIN DE LA CUENTA
        '***********************
        If Not Account_Database.MakeQuery("SELECT id, username, email, password, salt, gemas, status FROM cuentas WHERE UPPER(username) = (?)", False, UCase$(UserName)) Then
            Call WriteErrorMsg(UserIndex, "Error al cargar la cuenta.")
            Call CloseUser(UserIndex)
            Exit Sub
        
        End If
            
        'Guardo la información de la cuenta
        .AccountInfo.ID = CInt(Account_Database.Database_RecordSet!ID)
        .AccountInfo.UserName = Account_Database.Database_RecordSet!UserName
        .AccountInfo.Email = Account_Database.Database_RecordSet!Email
        .AccountInfo.Password = Account_Database.Database_RecordSet!Password
        .AccountInfo.Salt = Account_Database.Database_RecordSet!Salt
        .AccountInfo.Gemas = CLng(Account_Database.Database_RecordSet!Gemas)
        .AccountInfo.status = CBool(Account_Database.Database_RecordSet!status)
            
        Set Account_Database.Database_RecordSet = Nothing
            
        '***********************
        'OBTENEMOS TODOS LOS PJ
        '***********************

        .AccountInfo.NumPjs = 0
        
        If User_Database.MakeQuery("SELECT id, name, level, body_id, head_id, weapon_id, shield_id, helmet_id, race_id, class_id, pos_map, rep_average, is_dead FROM personaje WHERE cuenta_id = (?) AND deleted = FALSE", _
                                    False, UserList(UserIndex).AccountInfo.ID) Then
        

            User_Database.Database_RecordSet.MoveFirst
        
            While Not User_Database.Database_RecordSet.EOF
                
                'Incrementamos la cantidad de PJ creados actualmnente
                .AccountInfo.NumPjs = .AccountInfo.NumPjs + 1
    
                .AccountInfo.AccountPJ(.AccountInfo.NumPjs).ID = User_Database.Database_RecordSet!ID
                .AccountInfo.AccountPJ(.AccountInfo.NumPjs).Name = User_Database.Database_RecordSet!Name
                .AccountInfo.AccountPJ(.AccountInfo.NumPjs).body = User_Database.Database_RecordSet!body_id
                .AccountInfo.AccountPJ(.AccountInfo.NumPjs).Head = User_Database.Database_RecordSet!head_id
                .AccountInfo.AccountPJ(.AccountInfo.NumPjs).weapon = User_Database.Database_RecordSet!weapon_id
                .AccountInfo.AccountPJ(.AccountInfo.NumPjs).shield = User_Database.Database_RecordSet!shield_id
                .AccountInfo.AccountPJ(.AccountInfo.NumPjs).helmet = User_Database.Database_RecordSet!helmet_id
                .AccountInfo.AccountPJ(.AccountInfo.NumPjs).Class = User_Database.Database_RecordSet!class_id
                .AccountInfo.AccountPJ(.AccountInfo.NumPjs).race = User_Database.Database_RecordSet!race_id
                .AccountInfo.AccountPJ(.AccountInfo.NumPjs).Map = User_Database.Database_RecordSet!pos_map
                .AccountInfo.AccountPJ(.AccountInfo.NumPjs).level = User_Database.Database_RecordSet!level
                .AccountInfo.AccountPJ(.AccountInfo.NumPjs).criminal = (User_Database.Database_RecordSet!rep_average < 0)
                .AccountInfo.AccountPJ(.AccountInfo.NumPjs).dead = User_Database.Database_RecordSet!is_dead
                .AccountInfo.AccountPJ(.AccountInfo.NumPjs).gameMaster = EsGmChar(User_Database.Database_RecordSet!Name)
                    
                If .AccountInfo.AccountPJ(.AccountInfo.NumPjs).gameMaster = True Then TieneGM = True
                        
                User_Database.Database_RecordSet.MoveNext
            Wend
    
            Set User_Database.Database_RecordSet = Nothing
    
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
    
    #If DBConexionUnica = 0 Then
        Call Account_Database.Database_Close
        Call User_Database.Database_Close
    #End If
    
    Call WriteEnviarPJUserAccount(UserIndex)

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
    'Author: Lorwik
    'Last Modification: 06/04/2021
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    #If DBConexionUnica = 0 Then
        Call Account_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If Account_Database.CheckSQLStatus = False Then Account_Database.Database_Reconnect
    #End If

    If Not Account_Database.MakeQuery("SELECT id FROM cuentas WHERE UPPER(username) = (?)", False, UCase$(UserName)) Then
        CuentaExisteDatabase = False
        Exit Function

    End If

    CuentaExisteDatabase = (Account_Database.Database_RecordSet.RecordCount > 0)
    Set Account_Database.Database_RecordSet = Nothing
    
    #If DBConexionUnica = 0 Then
        Call Account_Database.Database_Close
    #End If

    Exit Function

ErrorHandler:
    If Err.Number = -1207576359 Then _
        Call Account_Database.Database_Reconnect

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
        Call Account_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If Account_Database.CheckSQLStatus = False Then Account_Database.Database_Reconnect
    #End If
    
    If Not Account_Database.MakeQuery("SELECT status FROM cuentas WHERE UPPER(username) = (?)", False, UCase$(UserName)) Then
       CuentaVerificada = False
        Exit Function

    End If

    CuentaVerificada = CBool(Account_Database.Database_RecordSet!status)
    Set Account_Database.Database_RecordSet = Nothing
    
    #If DBConexionUnica = 0 Then
        Call Account_Database.Database_Close
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
    'hacemos una consulta buscando el nombre del personaje y el cuenta_id de
    'la persona que quieres entrar al personajeque le pasamos, si obtenemos 1 resultado
    'el personaje pertenece a la cuenta
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    If Not User_Database.MakeQuery("SELECT id FROM personaje WHERE UPPER(name) = (?) AND cuenta_id = (?)", False, UCase$(UserName), UserList(UserIndex).AccountInfo.ID) Then
        PersonajePerteneceCuenta = False
        Exit Function

    End If

    PersonajePerteneceCuenta = (User_Database.Database_RecordSet.RecordCount > 0)
    Set Account_Database.Database_RecordSet = Nothing
    
    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
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
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If
    
    If Not User_Database.MakeQuery("SELECT COUNT(*) FROM personaje WHERE deleted = 0 and cuenta_id = (?)", False, UserList(UserIndex).AccountInfo.ID) Then
        GetCountUserAccount = 0
        Exit Function

    End If

    GetCountUserAccount = val(User_Database.Database_RecordSet.Fields(0).Value)
    Set Account_Database.Database_RecordSet = Nothing
    
    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
    #End If

    Exit Function
ErrorHandler:
    Call LogDatabaseError("Error in GetUserTrainingTimeDatabase: UserIndex: " & UserIndex & " - Hash: " & UserList(UserIndex).AccountInfo.ID & ". " & Err.Number & " - " & Err.description)

End Function

Public Sub BorrarUsuarioDatabase(ByVal UserName As String)

    '***************************************************
    'Author: Lorwik
    'Last Modification: 07/04/2021
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    Call User_Database.MakeQuery("UPDATE personaje SET name = (?), deleted = TRUE WHERE UPPER(name) = (?)", True, UCase$(UserName) & "_deleted", UCase$(UserName))
    
    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
    #End If

    Exit Sub

ErrorHandler:
    Call LogDatabaseError("Error in BorrarUsuarioDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Function GetAccountSalt(ByVal AccountName As String) As String

    '***************************************************
    'Author: Lorwik
    'Last Modification: 07/04/2021
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    #If DBConexionUnica = 0 Then
        Call Account_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If Account_Database.CheckSQLStatus = False Then Account_Database.Database_Reconnect
    #End If
    
    If Not Account_Database.MakeQuery("SELECT salt FROM cuentas WHERE UPPER(username) = (?)", False, UCase$(AccountName)) Then
        GetAccountSalt = vbNullString
        Exit Function

    End If

    GetAccountSalt = Account_Database.Database_RecordSet!Salt
    Set Account_Database.Database_RecordSet = Nothing
    
    #If DBConexionUnica = 0 Then
        Call Account_Database.Database_Close
    #End If

    Exit Function
ErrorHandler:
    Call LogDatabaseError("Error in GetAccountSalt: " & AccountName & ". " & Err.Number & " - " & Err.description)

End Function

Public Function GetAccountPassword(ByVal AccountName As String) As String

    '***************************************************
    'Author: Lorwik
    'Last Modification: 07/04/2021
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    #If DBConexionUnica = 0 Then
        Call Account_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If Account_Database.CheckSQLStatus = False Then Account_Database.Database_Reconnect
    #End If

    If Not Account_Database.MakeQuery("SELECT password FROM cuentas WHERE UPPER(username) = (?)", False, UCase$(AccountName)) Then
        GetAccountPassword = vbNullString
        Exit Function

    End If

    GetAccountPassword = Account_Database.Database_RecordSet!Password
    Set Account_Database.Database_RecordSet = Nothing
        
    #If DBConexionUnica = 0 Then
        Call Account_Database.Database_Close
    #End If

    Exit Function
ErrorHandler:
    Call LogDatabaseError("Error in GetAccountPassword: " & AccountName & ". " & Err.Number & " - " & Err.description)

End Function

Public Function GetAccountID(ByVal UserName As String) As Long

    '***************************************************
    'Author: Lorwik
    'Last Modification: 06/04/2021
    'Descripcion: Devuelve la ID de la cuenta del usuario solicitado
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    If Not User_Database.MakeQuery("SELECT cuenta_id FROM personaje WHERE UPPER(name) = (?)", False, UCase$(UserName)) Then
        GetAccountID = -1
        Exit Function

    End If

    GetAccountID = User_Database.Database_RecordSet!cuenta_id
    Set Account_Database.Database_RecordSet = Nothing
        
    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
    #End If

    Exit Function
    
ErrorHandler:
    Call LogDatabaseError("Error in GetAccountID: " & UserName & ". " & Err.Number & " - " & Err.description)

End Function

Public Function GetUserEmail(ByVal UserName As String) As String

    '***************************************************
    'Author: Lorwik
    'Last Modification: 07/04/2021
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    #If DBConexionUnica = 0 Then
        Call Account_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If Account_Database.CheckSQLStatus = False Then Account_Database.Database_Reconnect
    #End If

    If Not User_Database.MakeQuery("SELECT email FROM cuentas WHERE id = (?)", False, GetAccountID(UserName)) Then
        GetUserEmail = vbNullString
        Exit Function

    End If

    GetUserEmail = Account_Database.Database_RecordSet!UserName
    Set Account_Database.Database_RecordSet = Nothing
    
    #If DBConexionUnica = 0 Then
        Call Account_Database.Database_Close
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
    'Author: Lorwik
    'Last Modification: ????
    'Descripcion: Crea una nueva cuenta desde el server
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    'Si perdimos la conexion reconectamos
    If Account_Database.CheckSQLStatus = False Then Account_Database.Database_Reconnect

    query = "INSERT INTO cuentas SET username = (?), email = (?), password = (?), salt = (?), id_confirmacion = 'VERIFICADA', status = '1', date_created = NOW(), date_last_login = NOW();"

    Call Account_Database.MakeQuery(query, True, UserName, Email, Password, Salt)

    SaveNewAccount = True
    
    Exit Function
ErrorHandler:
    Call LogDatabaseError("Error in SaveNewAccountDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)
    SaveNewAccount = False

End Function

Public Function SaveAccountEditGemasDatabase(ByVal UserName As String, ByVal Gemas As Long) As Boolean

    '***************************************************
    'Author: Lorwik
    'Last Modification: 30/04/2020
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String
    Dim UserAccId As Long
    
    #If DBConexionUnica = 0 Then
        Call Account_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If
    
    UserAccId = GetAccountID(UserName)
    
    '¿Obtuvimos una ID nula?
    If UserAccId <> -1 Then
    
        Call Account_Database.MakeQuery("UPDATE cuentas SET gemas = (?) WHERE id = " & UserAccId, True, Gemas)
        
        SaveAccountEditGemasDatabase = True
        
    Else
        SaveAccountEditGemasDatabase = False
        
    End If

    #If DBConexionUnica = 0 Then
        Call Account_Database.Database_Close
    #End If

    Exit Function
ErrorHandler:
    Call LogDatabaseError("Error in SaveAccountEditGemasDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)
    SaveAccountEditGemasDatabase = False

End Function

Public Function SaveAccountSumaGemasDatabase(ByVal UserName As String, ByVal Gemas As Long) As Boolean

    '***************************************************
    'Author: Lorwik
    'Last Modification: 30/04/2020
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String
    Dim UserAccId As Long
    
    #If DBConexionUnica = 0 Then
        Call Account_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    UserAccId = GetAccountID(UserName)
    
    '¿Obtuvimos una ID nula?
    If UserAccId <> -1 Then
    
        Call Account_Database.MakeQuery("UPDATE cuentas SET gemas = gemas + (?) WHERE id = " & UserAccId, True, Gemas)
        
        SaveAccountSumaGemasDatabase = True
        
    Else
        SaveAccountSumaGemasDatabase = False
        
    End If

    #If DBConexionUnica = 0 Then
        Call Account_Database.Database_Close
    #End If

    Exit Function
ErrorHandler:
    Call LogDatabaseError("Error in SaveAccountSumaGemasDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)
    SaveAccountSumaGemasDatabase = False

End Function

Public Function SaveAccountRestaGemasDatabase(ByVal UserName As String, ByVal Gemas As Long) As Boolean

    '***************************************************
    'Author: Lorwik
    'Last Modification: 30/04/2020
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String
    Dim UserAccId As Long

    #If DBConexionUnica = 0 Then
        Call Account_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    UserAccId = GetAccountID(UserName)
    
    '¿Obtuvimos una ID nula?
    If UserAccId <> -1 Then
    
        Call Account_Database.MakeQuery("UPDATE cuentas SET gemas = gemas - (?) WHERE id = " & UserAccId, True, Gemas)
        
        SaveAccountRestaGemasDatabase = True
        
    Else
        SaveAccountRestaGemasDatabase = False
        
    End If

    #If DBConexionUnica = 0 Then
        Call Account_Database.Database_Close
    #End If

    Exit Function
ErrorHandler:
    Call LogDatabaseError("Error in SaveAccountRestaGemasDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)
    SaveAccountRestaGemasDatabase = False
    
End Function

Public Function GetGemasDatabase(ByVal UserName As String) As Long

    '***************************************************
    'Author: Lorwik
    'Last Modification: 30/04/2020
    '***************************************************
    On Error GoTo ErrorHandler

    Dim UserAccId As Long
    
    #If DBConexionUnica = 0 Then
        Call Account_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    UserAccId = GetAccountID(UserName)
    
    If Not Account_Database.MakeQuery("SELECT gemas FROM cuentas WHERE id = (?)", False, UserAccId) Then
        GetGemasDatabase = 0
        Exit Function

    End If

    GetGemasDatabase = CLng(Account_Database.Database_RecordSet!Gemas)
    Set Account_Database.Database_RecordSet = Nothing
    
    #If DBConexionUnica = 0 Then
        Call Account_Database.Database_Close
    #End If

    Exit Function

ErrorHandler:
    Call LogDatabaseError("Error in GetGemasDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Function

Public Sub SaveAccountLastLoginDatabase(ByVal UserName As String, ByVal UserIP As String)

    '***************************************************
    'Author: Lorwik
    'Last Modification: 07/04/2021
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String
    
    #If DBConexionUnica = 0 Then
        Call Account_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If Account_Database.CheckSQLStatus = False Then Account_Database.Database_Reconnect
    #End If

    query = "UPDATE cuentas SET date_last_login = NOW(), last_ip = (?) WHERE UPPER(username) = (?)"
    Call Account_Database.MakeQuery(query, True, UserIP, UCase$(UserName))

    #If DBConexionUnica = 0 Then
        Call Account_Database.Database_Close
    #End If

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in SaveAccountLastLoginDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

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
        
        '¿Posicion invalida?
        If Posicion <= 0 Then Exit Sub

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
        Call WriteEnviarPJUserAccount(UserIndex)
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

Public Sub DeletePJCuenta(ByVal UserIndex As Integer, ByVal Slot As Byte)
'****************************************************
'Autor: Lorwik
'Fecha: 04/11/2020
'Descripcion: elimina un personaje de la lista de la cuentay, y lo reordena
'****************************************************

    Dim Count As Byte
    Dim i As Byte
    
    With UserList(UserIndex).AccountInfo

        '¿El Slot es el ultimo?
        If Slot = .NumPjs Then
            Call ResetPJAccountSlot(UserIndex, Slot)
            
        Else
            
            'Primero borro el Slot
            Call ResetPJAccountSlot(UserIndex, Slot)
        
            For i = Slot To .NumPjs - 1
            
                .AccountPJ(i).ID = .AccountPJ(i + 1).ID
                .AccountPJ(i).Name = .AccountPJ(i + 1).Name
                .AccountPJ(i).body = .AccountPJ(i + 1).body
                .AccountPJ(i).Head = .AccountPJ(i + 1).Head
                .AccountPJ(i).weapon = .AccountPJ(i + 1).weapon
                .AccountPJ(i).shield = .AccountPJ(i + 1).shield
                .AccountPJ(i).helmet = .AccountPJ(i + 1).helmet
                .AccountPJ(i).Class = .AccountPJ(i + 1).Class
                .AccountPJ(i).race = .AccountPJ(i + 1).race
                .AccountPJ(i).Map = .AccountPJ(i + 1).Map
                .AccountPJ(i).level = .AccountPJ(i + 1).level
                .AccountPJ(i).criminal = .AccountPJ(i + 1).criminal
                .AccountPJ(i).dead = .AccountPJ(i + 1).dead
                .AccountPJ(i).gameMaster = .AccountPJ(i + 1).gameMaster
                
                'Voy limpiando
                Call ResetPJAccountSlot(UserIndex, i + 1)
            
            Next i
        
        End If

        'Restamos 1
        .NumPjs = .NumPjs - 1

    End With

    Call WriteEnviarPJUserAccount(UserIndex)
    
End Sub

Public Sub ResetPJAccountSlot(ByVal UserIndex As Integer, ByVal Slot As Byte)
'****************************************************
'Autor: Lorwik
'Fecha: 04/11/2020
'Descripcion: Limpia un slot de la cuenta
'****************************************************

    With UserList(UserIndex).AccountInfo
        .AccountPJ(Slot).ID = 0
        .AccountPJ(Slot).Name = vbNullString
        .AccountPJ(Slot).body = 0
        .AccountPJ(Slot).Head = 0
        .AccountPJ(Slot).weapon = 0
        .AccountPJ(Slot).shield = 0
        .AccountPJ(Slot).helmet = 0
        .AccountPJ(Slot).Class = 0
        .AccountPJ(Slot).race = 0
        .AccountPJ(Slot).Map = 0
        .AccountPJ(Slot).level = 0
        .AccountPJ(Slot).criminal = False
        .AccountPJ(Slot).dead = False
        .AccountPJ(Slot).gameMaster = False
        
    End With
            
End Sub


