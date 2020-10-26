Attribute VB_Name = "modDatabase"
'Argentum Online Libre
'Database connection module
'Obtained from GS-Zone
'Adapted and modified by Juan Andres Dalmasso (CHOTS)
'September 2018

Option Explicit

Public Database_DataSource As String
Public Database_Host       As String
Public Database_Name       As String
Public Database_Username   As String
Public Database_Password   As String
Public Database_Connection As ADODB.Connection
Public Database_RecordSet  As ADODB.Recordset
 
Public Sub Database_Connect()

    '************************************************************************************
    'Author: Juan Andres Dalmasso
    'Last Modification: 21/09/2019
    '21/09/2019 Jopi - Agregue soporte a conexion via DSN. Solo para usuarios avanzados.
    '************************************************************************************
    On Error GoTo ErrorHandler
 
    Set Database_Connection = New ADODB.Connection
    
    If Len(Database_DataSource) <> 0 Then
    
        Database_Connection.ConnectionString = "DATA SOURCE=" & Database_DataSource & ";"
        
    Else
    
        Database_Connection.ConnectionString = "DRIVER={MySQL ODBC 8.0 ANSI Driver};" & _
                                               "SERVER=" & Database_Host & ";" & _
                                               "DATABASE=" & Database_Name & ";" & _
                                               "USER=" & Database_Username & ";" & _
                                               "PASSWORD=" & Database_Password & ";" & _
                                               "OPTION=3"
    End If
    
    'Debug.Print Database_Connection.ConnectionString
    
    Database_Connection.CursorLocation = adUseClient
    Database_Connection.Open

    Exit Sub
    
ErrorHandler:
    Call LogDatabaseError("Database Error: " & Err.Number & " - " & Err.description)
    Debug.Print "Database Error: " & Err.Number & " - " & Err.description

End Sub

Public Sub Database_Close()

    '***************************************************
    'Author: Juan Andres Dalmasso
    'Last Modification: 18/09/2018
    '***************************************************
    On Error GoTo ErrorHandler
     
    Database_Connection.Close
    Set Database_Connection = Nothing
     
    Exit Sub
     
ErrorHandler:
    Call LogDatabaseError("Unable to close Mysql Database: " & Err.Number & " - " & Err.description)

End Sub

Public Function Database_Reconnect() As Boolean
'***************************************************
'Author: Lorwik
'Fecha: 13/09/2020
'Descripcion: Reconexión de la base de datos
'***************************************************
    Dim CerrolaConexion As Boolean
    
    'Si la conexión ya existia...
    If Database_Connection Is Nothing Then
        CerrolaConexion = False
        
    Else
        Call Database_Close
        CerrolaConexion = True
        
    End If
    
    Call Database_Connect
    
    Call LogDatabaseError("Base de datos reconectada. ¿Se cerro la conexion?: " & CerrolaConexion)
    
End Function

Public Function CheckSQLStatus() As Boolean
    '***************************************************
    'Author: Lorwik
    'Fecha: 17/07/2020
    'Descripcion: Comprobamos el estado de la conexion a la base de datos.
    '***************************************************
    
    If Database_Connection Is Nothing Then
        CheckSQLStatus = False
        Exit Function
    End If
    
    If Database_Connection.State = 0 Then
        CheckSQLStatus = False
        Exit Function
    End If
    
    'La conexion es correcta
    CheckSQLStatus = True

End Function

Sub SaveUserToDatabase(ByVal UserIndex As Integer, _
                       Optional ByVal SaveTimeOnline As Boolean = True)
    '*************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last modified: 14/10/2018
    'Saves the User to the database
    '*************************************************

    On Error GoTo ErrorHandler

    With UserList(UserIndex)

        If .ID > 0 Then
            Call UpdateUserToDatabase(UserIndex, SaveTimeOnline)
        Else
            Call InsertUserToDatabase(UserIndex, SaveTimeOnline)

        End If

    End With
    
    Exit Sub

ErrorHandler:
    Call LogDatabaseError("Unable to save User to Mysql Database: " & UserList(UserIndex).Name & ". " & Err.Number & " - " & Err.description)

End Sub

Sub InsertUserToDatabase(ByVal UserIndex As Integer, _
                         Optional ByVal SaveTimeOnline As Boolean = True)
    '*************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last modified: 04/10/2018
    'Inserts a new user to the database, then gets its ID and assigns it
    '*************************************************

    On Error GoTo ErrorHandler

    Dim query  As String

    Dim UserID As Integer

    Dim LoopC  As Byte

    #If DBConexionUnica = 0 Then
        Call Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If CheckSQLStatus = False Then Database_Reconnect
    #End If

    'Basic user data
    With UserList(UserIndex)
        query = "INSERT INTO usuario SET "
        query = query & "name = '" & .Name & "', "
        query = query & "account_id = " & .AccountInfo.ID & ", "
        query = query & "level = " & .Stats.ELV & ", "
        query = query & "exp = " & .Stats.Exp & ", "
        query = query & "elu = " & .Stats.ELU & ", "
        query = query & "genre_id = " & .Genero & ", "
        query = query & "race_id = " & .Raza & ", "
        query = query & "class_id = " & .clase & ", "
        query = query & "home_id = " & .Hogar & ", "
        query = query & "description = '" & .Desc & "', "
        query = query & "gold = " & .Stats.Gld & ", "
        query = query & "free_skillpoints = " & .Stats.SkillPts & ", "
        query = query & "assigned_skillpoints = " & .Counters.AsignedSkills & ", "
        query = query & "elo = " & .Stats.ELO & ", "
        query = query & "pos_map = " & .Pos.Map & ", "
        query = query & "pos_x = " & .Pos.X & ", "
        query = query & "pos_y = " & .Pos.Y & ", "
        query = query & "body_id = " & .Char.body & ", "
        query = query & "head_id = " & .Char.Head & ", "
        query = query & "weapon_id = " & .Char.WeaponAnim & ", "
        query = query & "helmet_id = " & .Char.CascoAnim & ", "
        query = query & "shield_id = " & .Char.ShieldAnim & ", "
        query = query & "items_amount = " & .Invent.NroItems & ", "
        query = query & "slot_armour = " & .Invent.ArmourEqpSlot & ", "
        query = query & "slot_weapon = " & .Invent.WeaponEqpSlot & ", "
        query = query & "min_hp = " & .Stats.MinHp & ", "
        query = query & "max_hp = " & .Stats.MaxHp & ", "
        query = query & "min_man = " & .Stats.MinMAN & ", "
        query = query & "max_man = " & .Stats.MaxMAN & ", "
        query = query & "min_sta = " & .Stats.MinSta & ", "
        query = query & "max_sta = " & .Stats.MaxSta & ", "
        query = query & "min_ham = " & .Stats.MinHam & ", "
        query = query & "max_ham = " & .Stats.MaxHam & ", "
        query = query & "min_sed = " & .Stats.MinAGU & ", "
        query = query & "max_sed = " & .Stats.MaxAGU & ", "
        query = query & "min_hit = " & .Stats.MinHIT & ", "
        query = query & "max_hit = " & .Stats.MaxHIT & ", "
        query = query & "rep_noble = " & .Reputacion.NobleRep & ", "
        query = query & "rep_plebe = " & .Reputacion.PlebeRep & ", "
        query = query & "rep_average = " & .Reputacion.Promedio & ","
        query = query & "profesionA = " & .Profesion(0).Profesion & ","
        query = query & "ProfesionB = " & .Profesion(1).Profesion & ";"

        'Insert the user
        Call Database_Connection.Execute(query)

        'Get the user ID
        Set Database_RecordSet = Database_Connection.Execute("SELECT LAST_INSERT_ID();")

        If Database_RecordSet.BOF Or Database_RecordSet.EOF Then
            UserID = 1

        End If

        UserID = val(Database_RecordSet.Fields(0).Value)
        Set Database_RecordSet = Nothing

        .ID = UserID

        'User attributes
        query = "INSERT INTO attribute (user_id, number, value) VALUES "

        For LoopC = 1 To NUMATRIBUTOS
            query = query & "("
            query = query & .ID & ", "
            query = query & LoopC & ", "
            query = query & .Stats.UserAtributos(LoopC) & ")"

            If LoopC < NUMATRIBUTOS Then
                query = query & ", "
            Else
                query = query & ";"

            End If

        Next LoopC

        Call Database_Connection.Execute(query)

        'User spells
        query = "INSERT INTO spell (user_id, number, spell_id) VALUES "

        For LoopC = 1 To MAXUSERHECHIZOS
            query = query & "("
            query = query & .ID & ", "
            query = query & LoopC & ", "
            query = query & .Stats.UserHechizos(LoopC) & ")"

            If LoopC < MAXUSERHECHIZOS Then
                query = query & ", "
            Else
                query = query & ";"

            End If

        Next LoopC

        Call Database_Connection.Execute(query)

        'User inventory
        query = "INSERT INTO inventory_item (user_id, "
        
        For LoopC = 1 To MAX_INVENTORY_SLOTS
            query = query & "item_id" & LoopC & ", amount" & LoopC & ", is_equipped" & LoopC
            If LoopC < MAX_INVENTORY_SLOTS Then query = query & ", "
        Next LoopC
        
        query = query & ") VALUES (" & .ID & ", "
        
        For LoopC = 1 To MAX_INVENTORY_SLOTS

            query = query & .Invent.Object(LoopC).ObjIndex & ", "
            query = query & .Invent.Object(LoopC).Amount & ", "
            query = query & .Invent.Object(LoopC).Equipped
            If LoopC < MAX_INVENTORY_SLOTS Then query = query & ", "
            
        Next LoopC
        
        query = query & ");"

        Call Database_Connection.Execute(query)

        'User Boveda
        query = "INSERT INTO bank_item (user_id, number, item_id, amount) VALUES "

        For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
            query = query & "("
            query = query & .ID & ", "
            query = query & LoopC & ", "
            query = query & .BancoInvent.Object(LoopC).ObjIndex & ", "
            query = query & .BancoInvent.Object(LoopC).Amount & ")"

            If LoopC < MAX_BANCOINVENTORY_SLOTS Then
                query = query & ", "
            Else
                query = query & ";"

            End If

        Next LoopC

        Call Database_Connection.Execute(query)

        'User skills
        query = "INSERT INTO skillpoint (user_id, number, value, exp, elu) VALUES "

        For LoopC = 1 To NUMSKILLS
            query = query & "("
            query = query & .ID & ", "
            query = query & LoopC & ", "
            query = query & .Stats.UserSkills(LoopC) & ", "
            query = query & .Stats.ExpSkills(LoopC) & ", "
            query = query & .Stats.EluSkills(LoopC) & ")"

            If LoopC < NUMSKILLS Then
                query = query & ", "
            Else
                query = query & ";"

            End If

        Next LoopC

        Call Database_Connection.Execute(query)
        
        'Quests
        query = "INSERT INTO quest (idquest, user_id, npcs, estado) VALUES "
        
        For LoopC = 1 To MAXQUESTS
        
            query = query & "("
            query = query & LoopC & ", "
            query = query & .ID & ", "
            query = query & "0, "
            query = query & "0)"
            
            If LoopC < MAXQUESTS Then
                query = query & ", "
            Else
                query = query & ";"

            End If

        Next LoopC
        
        Call Database_Connection.Execute(query)
        
        'User Profesion Primaria
        query = "INSERT INTO profesion_primaria (user_id, number, receta_id) VALUES "

        For LoopC = 1 To MAXUSERRECETAS
            query = query & "("
            query = query & .ID & ", "
            query = query & LoopC & ", "
            query = query & .Profesion(0).Recetas(LoopC) & ")"

            If LoopC < MAXUSERRECETAS Then
                query = query & ", "
            Else
                query = query & ";"

            End If

        Next LoopC

        Call Database_Connection.Execute(query)
        
        'User Profesion Secundaria
        query = "INSERT INTO profesion_secundaria (user_id, number, receta_id) VALUES "

        For LoopC = 1 To MAXUSERRECETAS
            query = query & "("
            query = query & .ID & ", "
            query = query & LoopC & ", "
            query = query & .Profesion(1).Recetas(LoopC) & ")"

            If LoopC < MAXUSERRECETAS Then
                query = query & ", "
            Else
                query = query & ";"

            End If

        Next LoopC

        Call Database_Connection.Execute(query)

    End With
    
    #If DBConexionUnica = 0 Then
        Call Database_Close
    #End If

    Exit Sub

ErrorHandler:
    Call LogDatabaseError("Unable to INSERT User to Mysql Database: " & UserList(UserIndex).Name & ". " & Err.Number & " - " & Err.description)

End Sub

Sub UpdateUserToDatabase(ByVal UserIndex As Integer, _
                         Optional ByVal SaveTimeOnline As Boolean = True)
    '*************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last modified: 04/10/2018
    'Updates an existing user in the database
    '*************************************************

    On Error GoTo ErrorHandler

    Dim query  As String

    Dim UserID As Integer

    Dim LoopC  As Integer

    #If DBConexionUnica = 0 Then
        Call Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If CheckSQLStatus = False Then Database_Reconnect
    #End If

    'Basic user data
    With UserList(UserIndex)
        query = "UPDATE usuario SET "
        query = query & "name = '" & .Name & "', "
        query = query & "level = " & .Stats.ELV & ", "
        query = query & "exp = " & .Stats.Exp & ", "
        query = query & "elu = " & .Stats.ELU & ", "
        query = query & "genre_id = " & .Genero & ", "
        query = query & "race_id = " & .Raza & ", "
        query = query & "class_id = " & .clase & ", "
        query = query & "home_id = " & .Hogar & ", "
        query = query & "description = '" & .Desc & "', "
        query = query & "gold = " & .Stats.Gld & ", "
        query = query & "bank_gold = " & .Stats.Banco & ", "
        query = query & "free_skillpoints = " & .Stats.SkillPts & ", "
        query = query & "assigned_skillpoints = " & .Counters.AsignedSkills & ", "
        query = query & "elo = " & .Stats.ELO & ", "
        query = query & "pet_amount = " & .NroMascotas & ", "
        query = query & "pos_map = " & .Pos.Map & ", "
        query = query & "pos_x = " & .Pos.X & ", "
        query = query & "pos_y = " & .Pos.Y & ", "
        query = query & "last_map = " & .flags.lastMap & ", "
        query = query & "body_id = " & .Char.body & ", "
        query = query & "head_id = " & .OrigChar.Head & ", "
        query = query & "weapon_id = " & .Char.WeaponAnim & ", "
        query = query & "helmet_id = " & .Char.CascoAnim & ", "
        query = query & "shield_id = " & .Char.ShieldAnim & ", "
        query = query & "aura_id = " & .Char.AuraAnim & ", "
        query = query & "aura_color = " & .Char.AuraColor & ", "
        query = query & "heading = " & .Char.Heading & ", "
        query = query & "items_amount = " & .Invent.NroItems & ", "
        query = query & "slot_armour = " & .Invent.ArmourEqpSlot & ", "
        query = query & "slot_weapon = " & .Invent.WeaponEqpSlot & ", "
        query = query & "slot_helmet = " & .Invent.CascoEqpSlot & ", "
        query = query & "slot_shield = " & .Invent.EscudoEqpSlot & ", "
        query = query & "slot_ammo = " & .Invent.MunicionEqpSlot & ", "
        query = query & "slot_ship = " & .Invent.BarcoSlot & ", "
        query = query & "slot_ring = " & .Invent.AnilloEqpSlot & ", "
        query = query & "slot_bag = " & .Invent.MochilaEqpSlot & ", "
        query = query & "min_hp = " & .Stats.MinHp & ", "
        query = query & "max_hp = " & .Stats.MaxHp & ", "
        query = query & "min_man = " & .Stats.MinMAN & ", "
        query = query & "max_man = " & .Stats.MaxMAN & ", "
        query = query & "min_sta = " & .Stats.MinSta & ", "
        query = query & "max_sta = " & .Stats.MaxSta & ", "
        query = query & "min_ham = " & .Stats.MinHam & ", "
        query = query & "max_ham = " & .Stats.MaxHam & ", "
        query = query & "min_sed = " & .Stats.MinAGU & ", "
        query = query & "max_sed = " & .Stats.MaxAGU & ", "
        query = query & "min_hit = " & .Stats.MinHIT & ", "
        query = query & "max_hit = " & .Stats.MaxHIT & ", "
        query = query & "killed_npcs = " & .Stats.NPCsMuertos & ", "
        query = query & "killed_users = " & .Stats.UsuariosMatados & ", "
        query = query & "rep_asesino = " & .Reputacion.AsesinoRep & ", "
        query = query & "rep_bandido = " & .Reputacion.BandidoRep & ", "
        query = query & "rep_burgues = " & .Reputacion.BurguesRep & ", "
        query = query & "rep_ladron = " & .Reputacion.LadronesRep & ", "
        query = query & "rep_noble = " & .Reputacion.NobleRep & ", "
        query = query & "rep_plebe = " & .Reputacion.PlebeRep & ", "
        query = query & "rep_average = " & .Reputacion.Promedio & ", "
        query = query & "is_naked = " & .flags.Desnudo & ", "
        query = query & "is_poisoned = " & .flags.Envenenado & ", "
        query = query & "is_incinerado = " & .flags.Incinerado & ", "
        query = query & "is_hidden = " & .flags.Escondido & ", "
        query = query & "is_hungry = " & .flags.Hambre & ", "
        query = query & "is_thirsty = " & .flags.Sed & ", "
        query = query & "is_ban = " & .flags.Ban & ", "
        query = query & "is_dead = " & .flags.Muerto & ", "
        query = query & "is_sailing = " & .flags.Navegando & ", "
        query = query & "is_paralyzed = " & .flags.Paralizado & ", "
        query = query & "counter_pena = " & .Counters.Pena & ", "
        query = query & "pertenece_consejo_real = " & (.flags.Privilegios And PlayerType.RoyalCouncil) & ", "
        query = query & "pertenece_consejo_caos = " & (.flags.Privilegios And PlayerType.ChaosCouncil) & ", "
        query = query & "pertenece_real = " & .Faccion.ArmadaReal & ", "
        query = query & "pertenece_caos = " & .Faccion.FuerzasCaos & ", "
        query = query & "ciudadanos_matados = " & .Faccion.CiudadanosMatados & ", "
        query = query & "criminales_matados = " & .Faccion.CriminalesMatados & ", "
        query = query & "recibio_armadura_real = " & .Faccion.RecibioArmaduraReal & ", "
        query = query & "recibio_armadura_caos = " & .Faccion.RecibioArmaduraCaos & ", "
        query = query & "recibio_exp_real = " & .Faccion.RecibioExpInicialReal & ", "
        query = query & "recibio_exp_caos = " & .Faccion.RecibioExpInicialCaos & ", "
        query = query & "recompensas_real = " & .Faccion.RecompensasReal & ", "
        query = query & "recompensas_caos = " & .Faccion.RecompensasCaos & ", "
        query = query & "reenlistadas = " & .Faccion.Reenlistadas & ", "
        query = query & "fecha_ingreso = " & IIf(.Faccion.FechaIngreso <> vbNullString, "'" & .Faccion.FechaIngreso & "'", "NULL") & ", "
        query = query & "nivel_ingreso = " & .Faccion.NivelIngreso & ", "
        query = query & "matados_ingreso = " & .Faccion.MatadosIngreso & ", "
        query = query & "siguiente_recompensa = " & .Faccion.NextRecompensa & ", "
        query = query & "guild_index = " & .GuildIndex & ", "
        query = query & "is_global = " & .flags.Global & ", "
        query = query & "profesionA = " & .Profesion(0).Profesion & ", "
        query = query & "profesionB = " & .Profesion(1).Profesion & ", "
        query = query & "modocombate = " & IIf(.flags.ModoCombate = True, "1", "0") & ", "
        query = query & "seguro = " & IIf(.flags.Seguro = True, "1", "0") & " "
        query = query & "WHERE id = " & .ID & ";"
        Call Database_Connection.Execute(query)

        '*******************************************************************
        'Hechizos
        '*******************************************************************
        For LoopC = 1 To MAXUSERHECHIZOS
            query = "UPDATE spell SET "
            query = query & "spell_id = '" & .Stats.UserHechizos(LoopC) & "' "
            query = query & "WHERE user_id = '" & .ID & "' AND number = '" & LoopC & "'"

            Call Database_Connection.Execute(query)
        Next LoopC

        '*******************************************************************
        'Inventario
        '*******************************************************************
        
        query = "UPDATE inventory_item SET "
        
        For LoopC = 1 To MAX_INVENTORY_SLOTS
            
            query = query & "item_id" & LoopC & " = '" & .Invent.Object(LoopC).ObjIndex & "', "
            query = query & "amount" & LoopC & " = '" & .Invent.Object(LoopC).Amount & "', "
            query = query & "is_equipped" & LoopC & " = '" & .Invent.Object(LoopC).Equipped & "'"
            
            If LoopC < MAX_INVENTORY_SLOTS Then query = query & ", "
            
        Next LoopC
        
        query = query & " WHERE user_id = '" & .ID & "'"
        
        Call Database_Connection.Execute(query)

        '*******************************************************************
        'Boveda
        '*******************************************************************
        For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
            query = "UPDATE bank_item SET "
            query = query & "item_id = '" & .BancoInvent.Object(LoopC).ObjIndex & "', "
            query = query & "amount = '" & .BancoInvent.Object(LoopC).Amount & "' "
            query = query & "WHERE user_id = '" & .ID & "' AND number = '" & LoopC & "'"
            
            Call Database_Connection.Execute(query)
        Next LoopC

        '*******************************************************************
        'Skills
        '*******************************************************************
        For LoopC = 1 To NUMSKILLS
            query = "UPDATE skillpoint SET "
            query = query & "value = '" & .Stats.UserSkills(LoopC) & "', "
            query = query & "exp = '" & .Stats.ExpSkills(LoopC) & "', "
            query = query & "elu = '" & .Stats.EluSkills(LoopC) & "' "
            query = query & "WHERE user_id = '" & .ID & "' AND number = '" & LoopC & "'"
            
            Call Database_Connection.Execute(query)
        Next LoopC

        '*******************************************************************
        'Recetas
        '*******************************************************************
        For LoopC = 1 To MAXUSERRECETAS
            query = "UPDATE profesion_primaria SET "
            query = query & "receta_id = '" & .Profesion(0).Recetas(LoopC) & "' "
            query = query & "WHERE user_id = '" & .ID & "' AND number = '" & LoopC & "'"
            
            Call Database_Connection.Execute(query)
        Next LoopC
        
        For LoopC = 1 To MAXUSERRECETAS
            query = "UPDATE profesion_secundaria SET "
            query = query & "receta_id = '" & .Profesion(1).Recetas(LoopC) & "' "
            query = query & "WHERE user_id = '" & .ID & "' AND number = '" & LoopC & "'"
            
            Call Database_Connection.Execute(query)
        Next LoopC
        
        '*******************************************************************
        'Mascotas
        '*******************************************************************
        Dim petType As Integer
        For LoopC = 1 To MAXMASCOTAS
            query = "UPDATE pet SET "

            'CHOTS | I got this logic from SaveUserToCharfile
            If .MascotasIndex(LoopC) > 0 Then
                If Npclist(.MascotasIndex(LoopC)).Contadores.TiempoExistencia = 0 Then
                    petType = .MascotasType(LoopC)
                Else
                    petType = 0

                End If

            Else
                petType = .MascotasType(LoopC)

            End If

            query = query & "pet_id = '" & petType & "' "
            query = query & "WHERE user_id = '" & .ID & "' AND number = '" & LoopC & "'"

            Call Database_Connection.Execute(query)
        Next LoopC

    End With

    #If DBConexionUnica = 0 Then
        Call Database_Close
    #End If

    Exit Sub

ErrorHandler:
    Call LogDatabaseError("Unable to UPDATE usuario to Mysql Database: " & UserList(UserIndex).Name & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub UpdateUserQuest(ByVal UserIndex As Integer)
'************************************************************
'Autor: Lorwik
'Fecha: 28/06/2020
'Descripción: Guarda las quest del usuario en la base de datos
'************************************************************

    Dim query  As String
    Dim LoopC  As Integer
    Dim j      As Integer
    Dim tmpst  As String
    
    #If DBConexionUnica = 0 Then
        Call Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If CheckSQLStatus = False Then Database_Reconnect
    #End If
    
    'Basic user data
    With UserList(UserIndex)

        For LoopC = 1 To MAXQUESTS
        
            query = "UPDATE quest SET "
            query = query & "estado = '" & CInt(.QuestStats.Quests(LoopC).QuestStatus) & "', "           'Estado de la quest
            
            tmpst = "npcs = '0'"
            
            If .QuestStats.Quests(LoopC).QuestStatus = eStatusQuest.EnCurso Then
            
                '¿La quest requiere matar NPC?
                If QuestList(LoopC).RequiredNPCs > 0 Then
                    tmpst = "npcs = '"
                    For j = 1 To QuestList(LoopC).RequiredNPCs
                    
                        tmpst = tmpst & val(.QuestStats.Quests(LoopC).NPCsKilled(j))   'Cuantos NPCs se ha matado de los que requeridos
                        
                        If Not j = QuestList(LoopC).RequiredNPCs Then tmpst = tmpst & "."
                    Next j
                    
                    tmpst = tmpst & "'"

                End If
                
            End If
            
            query = query & tmpst
            query = query & " WHERE user_id = '" & .ID & "' AND idquest = '" & LoopC & "'"

            Call Database_Connection.Execute(query)
        Next LoopC
    End With

    #If DBConexionUnica = 0 Then
        Call Database_Close
    #End If

    Exit Sub

ErrorHandler:
    Call LogDatabaseError("Unable to UPDATE usuario to Mysql Database: " & UserList(UserIndex).Name & ". " & Err.Number & " - " & Err.description)
End Sub

Sub LoadUserFromDatabase(ByVal UserIndex As Integer)
    '*************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last modified: 09/10/2018
    'Loads the user from the database
    '*************************************************

    On Error GoTo ErrorHandler

    Dim query As String

    Dim LoopC As Byte

    #If DBConexionUnica = 0 Then
        Call Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If CheckSQLStatus = False Then Database_Reconnect
    #End If

    'Basic user data
    With UserList(UserIndex)
        query = "SELECT *, DATE_FORMAT(fecha_ingreso, '%Y-%m-%d') as 'fecha_ingreso_format' FROM usuario WHERE UPPER(name) ='" & UCase$(.Name) & "';"
        Set Database_RecordSet = Database_Connection.Execute(query)

        If Database_RecordSet.BOF Or Database_RecordSet.EOF Then Exit Sub

        'Start setting data
        .ID = Database_RecordSet!ID
        .Name = Database_RecordSet!Name
        .Stats.ELV = Database_RecordSet!level
        .Stats.Exp = Database_RecordSet!Exp
        .Stats.ELU = Database_RecordSet!ELU
        .Genero = Database_RecordSet!genre_id
        .Raza = Database_RecordSet!race_id
        .clase = Database_RecordSet!class_id
        .Hogar = Database_RecordSet!home_id
        .Desc = Database_RecordSet!description
        .Stats.Gld = Database_RecordSet!Gold
        .Stats.Banco = Database_RecordSet!bank_gold
        .Stats.SkillPts = Database_RecordSet!free_skillpoints
        .Counters.AsignedSkills = Database_RecordSet!assigned_skillpoints
        .Stats.ELO = Database_RecordSet!ELO
        .NroMascotas = Database_RecordSet!pet_amount
        .Pos.Map = Database_RecordSet!pos_map
        .Pos.X = Database_RecordSet!pos_x
        .Pos.Y = Database_RecordSet!pos_y
        .flags.lastMap = Database_RecordSet!last_map
        .OrigChar.body = Database_RecordSet!body_id
        .OrigChar.Head = Database_RecordSet!head_id
        .OrigChar.WeaponAnim = Database_RecordSet!weapon_id
        .OrigChar.CascoAnim = Database_RecordSet!helmet_id
        .OrigChar.ShieldAnim = Database_RecordSet!shield_id
        .OrigChar.Heading = Database_RecordSet!Heading
        .OrigChar.AuraAnim = Database_RecordSet!Aura_id
        .OrigChar.AuraColor = Database_RecordSet!Aura_color
        .Invent.NroItems = Database_RecordSet!items_amount
        .Invent.ArmourEqpSlot = SanitizeNullValue(Database_RecordSet!slot_armour, 0)
        .Invent.WeaponEqpSlot = SanitizeNullValue(Database_RecordSet!slot_weapon, 0)
        .Invent.CascoEqpSlot = SanitizeNullValue(Database_RecordSet!slot_helmet, 0)
        .Invent.EscudoEqpSlot = SanitizeNullValue(Database_RecordSet!slot_shield, 0)
        .Invent.MunicionEqpSlot = SanitizeNullValue(Database_RecordSet!slot_ammo, 0)
        .Invent.BarcoSlot = SanitizeNullValue(Database_RecordSet!slot_ship, 0)
        .Invent.AnilloEqpSlot = SanitizeNullValue(Database_RecordSet!slot_ring, 0)
        .Invent.MochilaEqpSlot = SanitizeNullValue(Database_RecordSet!slot_bag, 0)
        .Stats.MinHp = Database_RecordSet!min_hp
        .Stats.MaxHp = Database_RecordSet!max_hp
        .Stats.MinMAN = Database_RecordSet!min_man
        .Stats.MaxMAN = Database_RecordSet!max_man
        .Stats.MinSta = Database_RecordSet!min_sta
        .Stats.MaxSta = Database_RecordSet!max_sta
        .Stats.MinHam = Database_RecordSet!min_ham
        .Stats.MaxHam = Database_RecordSet!max_ham
        .Stats.MinAGU = Database_RecordSet!min_sed
        .Stats.MaxAGU = Database_RecordSet!max_sed
        .Stats.MinHIT = Database_RecordSet!min_hit
        .Stats.MaxHIT = Database_RecordSet!max_hit
        .Stats.NPCsMuertos = Database_RecordSet!killed_npcs
        .Stats.UsuariosMatados = Database_RecordSet!killed_users
        .Reputacion.AsesinoRep = Database_RecordSet!rep_asesino
        .Reputacion.BandidoRep = Database_RecordSet!rep_bandido
        .Reputacion.BurguesRep = Database_RecordSet!rep_burgues
        .Reputacion.LadronesRep = Database_RecordSet!rep_ladron
        .Reputacion.NobleRep = Database_RecordSet!rep_noble
        .Reputacion.PlebeRep = Database_RecordSet!rep_plebe
        .Reputacion.Promedio = Database_RecordSet!rep_average
        .flags.Desnudo = Database_RecordSet!is_naked
        .flags.Envenenado = Database_RecordSet!is_poisoned
        .flags.Incinerado = Database_RecordSet!is_incinerado
        .flags.Escondido = Database_RecordSet!is_hidden
        .flags.Hambre = Database_RecordSet!is_hungry
        .flags.Sed = Database_RecordSet!is_thirsty
        .flags.Ban = Database_RecordSet!is_ban
        .flags.Muerto = Database_RecordSet!is_dead
        .flags.Navegando = Database_RecordSet!is_sailing
        .flags.Paralizado = Database_RecordSet!is_paralyzed
        .Counters.Pena = Database_RecordSet!counter_pena
        .flags.Global = Database_RecordSet!is_global
        .Profesion(0).Profesion = Database_RecordSet!ProfesionA
        .Profesion(1).Profesion = Database_RecordSet!ProfesionB
        .flags.ModoCombate = Database_RecordSet!ModoCombate
        .flags.Seguro = Database_RecordSet!Seguro
        
        If Database_RecordSet!pertenece_consejo_real Then
            .flags.Privilegios = .flags.Privilegios Or PlayerType.RoyalCouncil

        End If

        If Database_RecordSet!pertenece_consejo_caos Then
            .flags.Privilegios = .flags.Privilegios Or PlayerType.ChaosCouncil

        End If

        .Faccion.ArmadaReal = Database_RecordSet!pertenece_real
        .Faccion.FuerzasCaos = Database_RecordSet!pertenece_caos
        .Faccion.CiudadanosMatados = Database_RecordSet!ciudadanos_matados
        .Faccion.CriminalesMatados = Database_RecordSet!criminales_matados
        .Faccion.RecibioArmaduraReal = Database_RecordSet!recibio_armadura_real
        .Faccion.RecibioArmaduraCaos = Database_RecordSet!recibio_armadura_caos
        .Faccion.RecibioExpInicialReal = Database_RecordSet!recibio_exp_real
        .Faccion.RecibioExpInicialCaos = Database_RecordSet!recibio_exp_caos
        .Faccion.RecompensasReal = Database_RecordSet!recompensas_real
        .Faccion.RecompensasCaos = Database_RecordSet!recompensas_caos
        .Faccion.Reenlistadas = Database_RecordSet!Reenlistadas
        .Faccion.FechaIngreso = SanitizeNullValue(Database_RecordSet!fecha_ingreso_format, vbNullString)
        .Faccion.NivelIngreso = SanitizeNullValue(Database_RecordSet!nivel_ingreso, 0)
        .Faccion.MatadosIngreso = SanitizeNullValue(Database_RecordSet!matados_ingreso, 0)
        .Faccion.NextRecompensa = SanitizeNullValue(Database_RecordSet!siguiente_recompensa, 0)

        .GuildIndex = SanitizeNullValue(Database_RecordSet!Guild_Index, 0)

        Set Database_RecordSet = Nothing

        'User attributes
        query = "SELECT * FROM attribute WHERE user_id = " & .ID & ";"
        Set Database_RecordSet = Database_Connection.Execute(query)
    
        If Not Database_RecordSet.RecordCount = 0 Then
            Database_RecordSet.MoveFirst

            While Not Database_RecordSet.EOF

                .Stats.UserAtributos(Database_RecordSet!Number) = Database_RecordSet!Value
                .Stats.UserAtributosBackUP(Database_RecordSet!Number) = .Stats.UserAtributos(Database_RecordSet!Number)

                Database_RecordSet.MoveNext
            Wend

        End If

        Set Database_RecordSet = Nothing

        'User spells
        query = "SELECT * FROM spell WHERE user_id = " & .ID & ";"
        Set Database_RecordSet = Database_Connection.Execute(query)

        If Not Database_RecordSet.RecordCount = 0 Then
            Database_RecordSet.MoveFirst

            While Not Database_RecordSet.EOF

                .Stats.UserHechizos(Database_RecordSet!Number) = Database_RecordSet!spell_id

                Database_RecordSet.MoveNext
            Wend

        End If

        Set Database_RecordSet = Nothing

        'User pets
        query = "SELECT * FROM pet WHERE user_id = " & .ID & ";"
        Set Database_RecordSet = Database_Connection.Execute(query)

        If Not Database_RecordSet.RecordCount = 0 Then
            Database_RecordSet.MoveFirst

            While Not Database_RecordSet.EOF

                .MascotasType(Database_RecordSet!Number) = Database_RecordSet!pet_id

                Database_RecordSet.MoveNext
            Wend

        End If

        Set Database_RecordSet = Nothing

        'User inventory
        query = "SELECT * FROM inventory_item WHERE user_id = " & .ID & ";"
        Set Database_RecordSet = Database_Connection.Execute(query)

        If Not Database_RecordSet.RecordCount = 0 Then
            Database_RecordSet.MoveFirst
                
            For LoopC = 1 To MAX_INVENTORY_SLOTS
                .Invent.Object(LoopC).ObjIndex = Database_RecordSet("item_id" & LoopC)
                .Invent.Object(LoopC).Amount = Database_RecordSet("Amount" & LoopC)
                .Invent.Object(LoopC).Equipped = Database_RecordSet("is_equipped" & LoopC)
            Next LoopC
                
        End If

        Set Database_RecordSet = Nothing

        'User bank inventory
        query = "SELECT * FROM bank_item WHERE user_id = " & .ID & ";"
        Set Database_RecordSet = Database_Connection.Execute(query)

        If Not Database_RecordSet.RecordCount = 0 Then
            Database_RecordSet.MoveFirst

            While Not Database_RecordSet.EOF

                .BancoInvent.Object(Database_RecordSet!Number).ObjIndex = Database_RecordSet!item_id
                .BancoInvent.Object(Database_RecordSet!Number).Amount = Database_RecordSet!Amount

                Database_RecordSet.MoveNext
            Wend

        End If

        Set Database_RecordSet = Nothing

        'User skills
        query = "SELECT * FROM skillpoint WHERE user_id = " & .ID & ";"
        Set Database_RecordSet = Database_Connection.Execute(query)

        If Not Database_RecordSet.RecordCount = 0 Then
            Database_RecordSet.MoveFirst

            While Not Database_RecordSet.EOF

                .Stats.UserSkills(Database_RecordSet!Number) = Database_RecordSet!Value
                .Stats.ExpSkills(Database_RecordSet!Number) = Database_RecordSet!Exp
                .Stats.EluSkills(Database_RecordSet!Number) = Database_RecordSet!ELU

                Database_RecordSet.MoveNext
            Wend

        End If

        Set Database_RecordSet = Nothing
        
        'User Recetas Profesion Primaria
        query = "SELECT * FROM profesion_primaria WHERE user_id = " & .ID & ";"
        Set Database_RecordSet = Database_Connection.Execute(query)

        If Not Database_RecordSet.RecordCount = 0 Then
            Database_RecordSet.MoveFirst

            While Not Database_RecordSet.EOF

                .Profesion(0).Recetas(Database_RecordSet!Number) = Database_RecordSet!receta_id

                Database_RecordSet.MoveNext
            Wend

        End If

        Set Database_RecordSet = Nothing
        
        'User Recetas Profesion Secundaria
        query = "SELECT * FROM profesion_secundaria WHERE user_id = " & .ID & ";"
        Set Database_RecordSet = Database_Connection.Execute(query)

        If Not Database_RecordSet.RecordCount = 0 Then
            Database_RecordSet.MoveFirst

            While Not Database_RecordSet.EOF

                .Profesion(1).Recetas(Database_RecordSet!Number) = Database_RecordSet!receta_id

                Database_RecordSet.MoveNext
            Wend

        End If

        Set Database_RecordSet = Nothing

    End With

    #If DBConexionUnica = 0 Then
        Call Database_Close
    #End If

    Exit Sub

ErrorHandler:
    Call LogDatabaseError("Unable to LOAD User from Mysql Database: " & UserList(UserIndex).Name & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub LoadQuestStats(ByVal UserIndex As Integer)
    '*************************************************
    'Autor: Lorwik
    'Fecha: 23/06/2020
    'Carga las quest del usuario desde la base de datos
    '*************************************************

    On Error GoTo ErrorHandler

    Dim query       As String
    Dim Fields()    As String
    Dim Count       As Integer
    Dim j           As Integer
    Dim tmpint      As Integer
    
    #If DBConexionUnica = 0 Then
        Call Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If CheckSQLStatus = False Then Database_Reconnect
    #End If

    With UserList(UserIndex).QuestStats

        query = "SELECT * FROM quest WHERE user_id = '" & UserList(UserIndex).ID & "';"
        Set Database_RecordSet = Database_Connection.Execute(query)
    
        If Not Database_RecordSet.RecordCount = 0 Then
            Database_RecordSet.MoveFirst
            Count = 1
            
            While Not Database_RecordSet.EOF And Count <> NumQuests

                    '¿La quest requiere matar NPC?
                    If QuestList(Count).RequiredNPCs Then
                        ReDim .Quests(Count).NPCsKilled(1 To QuestList(Count).RequiredNPCs)
            
                        Fields = Split(Database_RecordSet!NPCs, ".")
            
                        For j = 1 To QuestList(Count).RequiredNPCs

                            If UBound(Fields()) > 0 Then
                                .Quests(Count).NPCsKilled(j) = CInt(Fields(j - 1))
                                Debug.Print j & " - " & .Quests(Count).NPCsKilled(j)
                            Else
                                    .Quests(Count).NPCsKilled(j) = 0
                            End If
                        Next j
     
                    .Quests(Count).QuestStatus = CByte(Database_RecordSet!estado)
                             
                    'Si la quest actual se termino, lo sumamos al contador de terminados
                    If .Quests(Count).QuestStatus = eStatusQuest.Terminada Then .NumQuestsDone = .NumQuestsDone + 1

                    Count = Count + 1
                        
                    Database_RecordSet.MoveNext
                        
                End If
            Wend
                
           Call ListarQuestsenCurso(UserIndex)
                
        End If

    End With
    
    Set Database_RecordSet = Nothing

    #If DBConexionUnica = 0 Then
        Call Database_Close
    #End If
    
    Exit Sub

ErrorHandler:
    Call LogDatabaseError("Unable to LOAD User from Mysql Database: " & UserList(UserIndex).Name & ". " & Err.Number & " - " & Err.description)

End Sub

Public Function PersonajeExisteDatabase(ByVal UserName As String) As Boolean
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

    query = "SELECT id FROM usuario WHERE UPPER(name) = '" & UCase$(UserName) & "' AND deleted = FALSE;"

    Set Database_RecordSet = Database_Connection.Execute(query)

    If Database_RecordSet.BOF Or Database_RecordSet.EOF Then
        PersonajeExisteDatabase = False
        Exit Function

    End If

    PersonajeExisteDatabase = (Database_RecordSet.RecordCount > 0)
    Set Database_RecordSet = Nothing
    
    #If DBConexionUnica = 0 Then
        Call Database_Close
    #End If

    Exit Function

ErrorHandler:
    Call LogDatabaseError("Error in PersonajeExisteDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Function

Public Function BANCheckDatabase(ByVal UserName As String) As Boolean

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 09/10/2018
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    #If DBConexionUnica = 0 Then
        Call Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If CheckSQLStatus = False Then Database_Reconnect
    #End If

    query = "SELECT is_ban FROM usuario WHERE UPPER(name) = '" & UCase$(UserName) & "';"

    Set Database_RecordSet = Database_Connection.Execute(query)

    If Database_RecordSet.BOF Or Database_RecordSet.EOF Then
        BANCheckDatabase = False
        Exit Function

    End If

    BANCheckDatabase = CBool(Database_RecordSet!is_ban)

    Set Database_RecordSet = Nothing
    
    #If DBConexionUnica = 0 Then
        Call Database_Close
    #End If

    Exit Function

ErrorHandler:
    Call LogDatabaseError("Error in BANCheckDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Function

Public Sub UnBanDatabase(ByVal UserName As String)

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

    query = "UPDATE usuario SET is_ban = FALSE WHERE UPPER(name) = '" & UCase$(UserName) & "';"

    Database_Connection.Execute (query)

    #If DBConexionUnica = 0 Then
        Call Database_Close
    #End If

    Exit Sub

ErrorHandler:
    Call LogDatabaseError("Error in UnBanDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Function GetUserGuildIndexDatabase(ByVal UserName As String) As Integer

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 09/10/2018
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String
    
    #If DBConexionUnica = 0 Then
        Call Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If CheckSQLStatus = False Then Database_Reconnect
    #End If

    query = "SELECT guild_index FROM usuario WHERE UPPER(name) = '" & UCase$(UserName) & "';"

    Set Database_RecordSet = Database_Connection.Execute(query)

    If Database_RecordSet.BOF Or Database_RecordSet.EOF Then
        GetUserGuildIndexDatabase = 0
        Exit Function

    End If

    GetUserGuildIndexDatabase = SanitizeNullValue(Database_RecordSet!Guild_Index, 0)
    Set Database_RecordSet = Nothing
    
    #If DBConexionUnica = 0 Then
        Call Database_Close
    #End If

    Exit Function

ErrorHandler:
    Call LogDatabaseError("Error in GetUserGuildIndexDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Function

Public Sub CopyUserDatabase(ByVal UserName As String, ByVal newName As String)

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

    query = "UPDATE usuario SET name = '" & UCase$(newName) & "' WHERE UPPER(name) = '" & UCase$(UserName) & "';"

    Database_Connection.Execute (query)

    #If DBConexionUnica = 0 Then
        Call Database_Close
    #End If

    Exit Sub

ErrorHandler:
    Call LogDatabaseError("Error in CopyUserDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub MarcarPjComoQueYaVotoDatabase(ByVal UserIndex As Integer, _
                                         ByVal NumeroEncuesta As Integer)

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

    query = "UPDATE usuario SET votes_amount = " & NumeroEncuesta & " WHERE id = " & UserList(UserIndex).ID & ";"

    Database_Connection.Execute (query)
    
    #If DBConexionUnica = 0 Then
        Call Database_Close
    #End If

    Exit Sub

ErrorHandler:
    Call LogDatabaseError("Error in MarcarPjComoQueYaVotoDatabase: " & UserList(UserIndex).Name & ". " & Err.Number & " - " & Err.description)

End Sub

Public Function PersonajeCantidadVotosDatabase(ByVal UserName As String) As Integer

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

    query = "SELECT votes_amount FROM usuario WHERE UPPER(name) = '" & UCase$(UserName) & "';"

    Set Database_RecordSet = Database_Connection.Execute(query)

    If Database_RecordSet.BOF Or Database_RecordSet.EOF Then
        PersonajeCantidadVotosDatabase = 0
        Exit Function

    End If

    PersonajeCantidadVotosDatabase = CInt(Database_RecordSet!votes_amount)
    Set Database_RecordSet = Nothing
    
    #If DBConexionUnica = 0 Then
        Call Database_Close
    #End If

    Exit Function

ErrorHandler:
    Call LogDatabaseError("Error in PersonajeCantidadVotosDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Function

Public Sub SaveBan(ByVal UserName As String, _
                           ByVal Reason As String, _
                           ByVal BannedBy As String)

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 10/10/2018
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query     As String

    Dim cantPenas As Byte

    cantPenas = GetUserAmountOfPunishments(UserName)

    #If DBConexionUnica = 0 Then
        Call Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If CheckSQLStatus = False Then Database_Reconnect
    #End If

    query = "UPDATE usuario SET is_ban = TRUE WHERE UPPER(name) = '" & UCase$(UserName) & "';"

    Database_Connection.Execute (query)

    query = "INSERT INTO punishment SET "
    query = query & "user_id = (SELECT id FROM usuario WHERE UPPER(name) = '" & UCase$(UserName) & "'), "
    query = query & "number = " & (cantPenas + 1) & ", "
    query = query & "reason = '" & BannedBy & ": BAN POR " & LCase$(Reason) & " " & Date & " " & time & "';"

    Database_Connection.Execute (query)

    #If DBConexionUnica = 0 Then
        Call Database_Close
    #End If

    Exit Sub

ErrorHandler:
    Call LogDatabaseError("Error in SaveBan: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Function GetUserAmountOfPunishments(ByVal UserName As String) As Integer

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

    query = "SELECT COUNT(1) as punishments FROM punishment WHERE user_id = (SELECT id FROM usuario WHERE UPPER(name) = '" & UCase$(UserName) & "')"

    Set Database_RecordSet = Database_Connection.Execute(query)

    If Database_RecordSet.BOF Or Database_RecordSet.EOF Then
        GetUserAmountOfPunishments = 0
        Exit Function

    End If

    GetUserAmountOfPunishments = CInt(Database_RecordSet!punishments)
    Set Database_RecordSet = Nothing
    
    #If DBConexionUnica = 0 Then
        Call Database_Close
    #End If

    Exit Function
ErrorHandler:
    Call LogDatabaseError("Error in GetUserAmountOfPunishments: " & UserName & ". " & Err.Number & " - " & Err.description)

End Function

Public Sub SendUserPunishments(ByVal UserIndex As Integer, _
                                       ByVal UserName As String, _
                                       ByVal Count As Integer)

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

    query = "SELECT * FROM punishment WHERE user_id = (SELECT id FROM usuario WHERE UPPER(name) = '" & UCase$(UserName) & "');"

    Set Database_RecordSet = Database_Connection.Execute(query)

    If Not Database_RecordSet.RecordCount = 0 Then
        Database_RecordSet.MoveFirst

        While Not Database_RecordSet.EOF

            Call WriteConsoleMsg(UserIndex, Database_RecordSet!Number & " - " & Database_RecordSet!Reason, FontTypeNames.FONTTYPE_INFO)

            Database_RecordSet.MoveNext
        Wend

    End If

    Set Database_RecordSet = Nothing
    
    #If DBConexionUnica = 0 Then
        Call Database_Close
    #End If

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in SendUserPunishments: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Function GetUserPos(ByVal UserName As String) As String

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

    query = "SELECT pos_map, pos_x, pos_y FROM usuario WHERE UPPER(name) = '" & UCase$(UserName) & "';"

    Set Database_RecordSet = Database_Connection.Execute(query)

    If Database_RecordSet.BOF Or Database_RecordSet.EOF Then
        GetUserPos = vbNullString
        Exit Function

    End If

    GetUserPos = Database_RecordSet!pos_map & "-" & Database_RecordSet!pos_x & "-" & Database_RecordSet!pos_y
    Set Database_RecordSet = Nothing
    
    #If DBConexionUnica = 0 Then
        Call Database_Close
    #End If

    Exit Function
ErrorHandler:
    Call LogDatabaseError("Error in GetUserPos: " & UserName & ". " & Err.Number & " - " & Err.description)

End Function

Public Sub SaveUserPunishment(ByVal UserName As String, _
                                      ByVal Number As Integer, _
                                      ByVal Reason As String)

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

    query = "INSERT INTO punishment SET "
    query = query & "user_id = (SELECT id FROM usuario WHERE UPPER(name) = '" & UCase$(UserName) & "'), "
    query = query & "number = " & Number & ", "
    query = query & "reason = '" & Reason & "';"

    Database_Connection.Execute (query)

    #If DBConexionUnica = 0 Then
        Call Database_Close
    #End If

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in SaveUserPunishment: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub AlterUserPunishment(ByVal UserName As String, _
                                       ByVal Number As Integer, _
                                       ByVal Reason As String)

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

    query = "UPDATE punishment SET "
    query = query & "reason = '" & Reason & "' "
    query = query & "WHERE number = " & Number & " AND user_id = (SELECT id FROM usuario WHERE UPPER(name) = '" & UCase$(UserName) & "');"

    Database_Connection.Execute (query)

    #If DBConexionUnica = 0 Then
        Call Database_Close
    #End If

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in AlterUserPunishment: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub ResetUserFacciones(ByVal UserName As String)

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

    query = "UPDATE usuario SET "
    query = query & "pertenece_real = FALSE, "
    query = query & "pertenece_caos = FALSE, "
    query = query & "ciudadanos_matados = 0, "
    query = query & "criminales_matados = FALSE, "
    query = query & "recibio_armadura_real = FALSE, "
    query = query & "recibio_armadura_caos = FALSE, "
    query = query & "recibio_exp_real = FALSE, "
    query = query & "recibio_exp_caos = FALSE, "
    query = query & "recompensas_real = 0, "
    query = query & "recompensas_caos = 0, "
    query = query & "reenlistadas = 0, "
    query = query & "fecha_ingreso = NULL, "
    query = query & "nivel_ingreso = NULL, "
    query = query & "matados_ingreso = NULL, "
    query = query & "siguiente_recompensa = NULL "
    query = query & "WHERE UPPER(name) = '" & UCase$(UserName) & "';"

    Database_Connection.Execute (query)

    #If DBConexionUnica = 0 Then
        Call Database_Close
    #End If

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in ResetUserFacciones: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub KickUserCouncils(ByVal UserName As String)

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

    query = "UPDATE usuario SET "
    query = query & "pertenece_consejo_real = FALSE, "
    query = query & "pertenece_consejo_caos = FALSE "
    query = query & "WHERE UPPER(name) = '" & UCase$(UserName) & "';"

    Database_Connection.Execute (query)

    #If DBConexionUnica = 0 Then
        Call Database_Close
    #End If

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in KickUserCouncils: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub KickUserFacciones(ByVal UserName As String)

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

    query = "UPDATE usuario SET "
    query = query & "pertenece_real = FALSE, "
    query = query & "pertenece_caos = FALSE "
    query = query & "WHERE UPPER(name) = '" & UCase$(UserName) & "';"

    Database_Connection.Execute (query)
    
    #If DBConexionUnica = 0 Then
        Call Database_Close
    #End If

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in KickUserFacciones: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub KickUserChaosLegion(ByVal UserName As String)

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

    query = "UPDATE usuario SET "
    query = query & "pertenece_caos = FALSE, "
    query = query & "reenlistadas = 200 "
    query = query & "WHERE UPPER(name) = '" & UCase$(UserName) & "';"

    Database_Connection.Execute (query)

    #If DBConexionUnica = 0 Then
        Call Database_Close
    #End If

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in KickUserChaosLegion: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub KickUserRoyalArmy(ByVal UserName As String)

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

    query = "UPDATE usuario SET "
    query = query & "pertenece_real = FALSE, "
    query = query & "reenlistadas = 200 "
    query = query & "WHERE UPPER(name) = '" & UCase$(UserName) & "';"

    Database_Connection.Execute (query)

    #If DBConexionUnica = 0 Then
        Call Database_Close
    #End If

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in KickUserRoyalArmy: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub UpdateUserLogged(ByVal UserName As String, ByVal Logged As Byte)

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

    query = "UPDATE usuario SET "
    query = query & "is_logged = " & IIf(Logged = 1, "TRUE", "FALSE") & " "
    query = query & "WHERE UPPER(name) = '" & UCase$(UserName) & "';"

    Database_Connection.Execute (query)

    #If DBConexionUnica = 0 Then
        Call Database_Close
    #End If

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in UpdateUserLogged: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Function GetUserLastIps(ByVal UserName As String) As String

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

    query = "SELECT last_ip FROM account WHERE id = (SELECT account_id FROM usuario WHERE UPPER(name) = '" & UCase$(UserName) & "');"

    Set Database_RecordSet = Database_Connection.Execute(query)

    If Database_RecordSet.BOF Or Database_RecordSet.EOF Then
        GetUserLastIps = vbNullString
        Exit Function

    End If

    GetUserLastIps = Database_RecordSet!last_ip
    Set Database_RecordSet = Nothing
    
    #If DBConexionUnica = 0 Then
        Call Database_Close
    #End If

    Exit Function
ErrorHandler:
    Call LogDatabaseError("Error in GetUserLastIps: " & UserName & ". " & Err.Number & " - " & Err.description)

End Function

Public Function GetUserSkills(ByVal UserName As String) As String

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 10/10/2018
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    GetUserSkills = vbNullString

    #If DBConexionUnica = 0 Then
        Call Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If CheckSQLStatus = False Then Database_Reconnect
    #End If

    query = "SELECT number, value FROM skillpoint WHERE user_id = (SELECT id FROM usuario WHERE UPPER(name) = '" & UCase$(UserName) & "');"

    Set Database_RecordSet = Database_Connection.Execute(query)

    If Not Database_RecordSet.RecordCount = 0 Then
        Database_RecordSet.MoveFirst

        While Not Database_RecordSet.EOF

            GetUserSkills = GetUserSkills & "CHAR>" & SkillsNames(Database_RecordSet!Number) & " = " & Database_RecordSet!Value & vbCrLf

            Database_RecordSet.MoveNext
        Wend

    End If

    Set Database_RecordSet = Nothing

    #If DBConexionUnica = 0 Then
        Call Database_Close
    #End If

    Exit Function
ErrorHandler:
    Call LogDatabaseError("Error in GetUserSkills: " & UserName & ". " & Err.Number & " - " & Err.description)

End Function

Public Function GetUserFreeSkills(ByVal UserName As String) As Integer

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

    query = "SELECT free_skillpoints FROM usuario WHERE UPPER(name) = '" & UCase$(UserName) & "';"

    Set Database_RecordSet = Database_Connection.Execute(query)

    If Database_RecordSet.BOF Or Database_RecordSet.EOF Then
        GetUserFreeSkills = 0
        Exit Function

    End If

    GetUserFreeSkills = CInt(Database_RecordSet!free_skillpoints)
    Set Database_RecordSet = Nothing
    
    #If DBConexionUnica = 0 Then
        Call Database_Close
    #End If

    Exit Function
ErrorHandler:
    Call LogDatabaseError("Error in GetUserFreeSkills: " & UserName & ". " & Err.Number & " - " & Err.description)

End Function

Public Sub SaveUserTrainingTime(ByVal UserName As String, _
                                        ByVal trainingTime As Long)

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

    query = "UPDATE usuario SET "
    query = query & "counter_training = " & trainingTime & " "
    query = query & "WHERE UPPER(name) = '" & UCase$(UserName) & "';"

    Database_Connection.Execute (query)

    #If DBConexionUnica = 0 Then
        Call Database_Close
    #End If

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in SaveUserTrainingTime: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Function GetUserTrainingTime(ByVal UserName As String) As Long

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

    query = "SELECT counter_training FROM usuario WHERE UPPER(name) = '" & UCase$(UserName) & "';"

    Set Database_RecordSet = Database_Connection.Execute(query)

    If Database_RecordSet.BOF Or Database_RecordSet.EOF Then
        GetUserTrainingTime = 0
        Exit Function

    End If

    GetUserTrainingTime = CLng(Database_RecordSet!counter_training)
    Set Database_RecordSet = Nothing
    
    #If DBConexionUnica = 0 Then
        Call Database_Close
    #End If

    Exit Function
ErrorHandler:
    Call LogDatabaseError("Error in GetUserTrainingTime: " & UserName & ". " & Err.Number & " - " & Err.description)

End Function

Public Function UserBelongsToRoyalArmy(ByVal UserName As String) As Boolean

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

    query = "SELECT pertenece_real FROM usuario WHERE UPPER(name) = '" & UCase$(UserName) & "' AND deleted = FALSE;"

    Set Database_RecordSet = Database_Connection.Execute(query)

    If Database_RecordSet.BOF Or Database_RecordSet.EOF Then
        UserBelongsToRoyalArmy = False
        Exit Function

    End If

    UserBelongsToRoyalArmy = CBool(Database_RecordSet!pertenece_real)
    Set Database_RecordSet = Nothing
    
    #If DBConexionUnica = 0 Then
        Call Database_Close
    #End If

    Exit Function

ErrorHandler:
    Call LogDatabaseError("Error in UserBelongsToRoyalArmy: " & UserName & ". " & Err.Number & " - " & Err.description)

End Function

Public Function UserBelongsToChaosLegion(ByVal UserName As String) As Boolean

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

    query = "SELECT pertenece_caos FROM usuario WHERE UPPER(name) = '" & UCase$(UserName) & "' AND deleted = FALSE;"

    Set Database_RecordSet = Database_Connection.Execute(query)

    If Database_RecordSet.BOF Or Database_RecordSet.EOF Then
        UserBelongsToChaosLegion = False
        Exit Function

    End If

    UserBelongsToChaosLegion = CBool(Database_RecordSet!pertenece_caos)
    Set Database_RecordSet = Nothing
    
    #If DBConexionUnica = 0 Then
        Call Database_Close
    #End If

    Exit Function

ErrorHandler:
    Call LogDatabaseError("Error in UserBelongsToChaosLegion: " & UserName & ". " & Err.Number & " - " & Err.description)

End Function

Public Function GetUserLevel(ByVal UserName As String) As Byte

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 09/10/2018
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    #If DBConexionUnica = 0 Then
        Call Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If CheckSQLStatus = False Then Database_Reconnect
    #End If

    query = "SELECT level FROM usuario WHERE UPPER(name) = '" & UCase$(UserName) & "';"

    Set Database_RecordSet = Database_Connection.Execute(query)

    If Database_RecordSet.BOF Or Database_RecordSet.EOF Then
        GetUserLevel = 0
        Exit Function

    End If

    GetUserLevel = CByte(Database_RecordSet!level)
    Set Database_RecordSet = Nothing
    
    #If DBConexionUnica = 0 Then
        Call Database_Close
    #End If

    Exit Function

ErrorHandler:
    Call LogDatabaseError("Error in GetUserLevel: " & UserName & ". " & Err.Number & " - " & Err.description)

End Function

Public Function GetUserPromedio(ByVal UserName As String) As Long

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 09/10/2018
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    #If DBConexionUnica = 0 Then
        Call Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If CheckSQLStatus = False Then Database_Reconnect
    #End If

    query = "SELECT rep_average FROM usuario WHERE UPPER(name) = '" & UCase$(UserName) & "';"

    Set Database_RecordSet = Database_Connection.Execute(query)

    If Database_RecordSet.BOF Or Database_RecordSet.EOF Then
        GetUserPromedio = 0
        Exit Function

    End If

    GetUserPromedio = CLng(Database_RecordSet!rep_average)
    Set Database_RecordSet = Nothing

    #If DBConexionUnica = 0 Then
        Call Database_Close
    #End If

    Exit Function

ErrorHandler:
    Call LogDatabaseError("Error in GetUserPromedio: " & UserName & ". " & Err.Number & " - " & Err.description)

End Function

Public Function GetUserReenlists(ByVal UserName As String) As Byte

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 09/10/2018
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    #If DBConexionUnica = 0 Then
        Call Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If CheckSQLStatus = False Then Database_Reconnect
    #End If

    query = "SELECT reenlistadas FROM usuario WHERE UPPER(name) = '" & UCase$(UserName) & "';"

    Set Database_RecordSet = Database_Connection.Execute(query)

    If Database_RecordSet.BOF Or Database_RecordSet.EOF Then
        GetUserReenlists = 0
        Exit Function

    End If

    GetUserReenlists = CByte(Database_RecordSet!Reenlistadas)
    Set Database_RecordSet = Nothing
    
    #If DBConexionUnica = 0 Then
        Call Database_Close
    #End If

    Exit Function

ErrorHandler:
    Call LogDatabaseError("Error in GetUserReenlists: " & UserName & ". " & Err.Number & " - " & Err.description)

End Function

Public Sub SaveUserReenlists(ByVal UserName As String, ByVal Reenlists As Byte)

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

    query = "UPDATE usuario SET "
    query = query & "reenlistadas = " & Reenlists & " "
    query = query & "WHERE UPPER(name) = '" & UCase$(UserName) & "';"

    Database_Connection.Execute (query)

    #If DBConexionUnica = 0 Then
        Call Database_Close
    #End If

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in SaveUserReenlists: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub SendUserStatsTxtDatabase(ByVal sendIndex As Integer, ByVal UserName As String)

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 30/10/2018
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    If Not PersonajeExiste(UserName) Then
        Call WriteConsoleMsg(sendIndex, "Pj Inexistente", FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(sendIndex, "Estadisticas de: " & UserName, FontTypeNames.FONTTYPE_INFO)

    #If DBConexionUnica = 0 Then
        Call Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If CheckSQLStatus = False Then Database_Reconnect
    #End If
    
        query = "SELECT level, exp, elu, min_sta, max_sta, min_hp, max_hp, min_man, max_man, min_hit, max_hit, gold FROM usuario WHERE UPPER(name) = '" & UCase$(UserName) & "';"

        Set Database_RecordSet = Database_Connection.Execute(query)

        If Database_RecordSet.BOF Or Database_RecordSet.EOF Then
            Call WriteConsoleMsg(sendIndex, "Pj Inexistente", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        Call WriteConsoleMsg(sendIndex, "Nivel: " & Database_RecordSet!level & "  EXP: " & Database_RecordSet!Exp & "/" & Database_RecordSet!ELU, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Energia: " & Database_RecordSet!min_sta & "/" & Database_RecordSet!max_sta, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Salud: " & Database_RecordSet!min_hp & "/" & Database_RecordSet!max_hp, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Mana: " & Database_RecordSet!min_man & "/" & Database_RecordSet!max_man, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Golpe: " & Database_RecordSet!min_hit & "/" & Database_RecordSet!max_hit, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Oro: " & Database_RecordSet!Gold, FontTypeNames.FONTTYPE_INFO)

        Set Database_RecordSet = Nothing
        
        #If DBConexionUnica = 0 Then
            Call Database_Close
        #End If

    End If

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in SendUserStatsTxtDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub SendUserMiniStatsTxtFromDatabase(ByVal sendIndex As Integer, _
                                            ByVal UserName As String)

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 11/10/2018
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    If Not PersonajeExiste(UserName) Then
        Call WriteConsoleMsg(sendIndex, "Pj Inexistente", FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(sendIndex, "Estadisticas de: " & UserName, FontTypeNames.FONTTYPE_INFO)

    #If DBConexionUnica = 0 Then
        Call Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If CheckSQLStatus = False Then Database_Reconnect
    #End If
    
        query = "SELECT killed_npcs, killed_users, ciudadanos_matados, criminales_matados, class_id, genre_id, race_id FROM usuario WHERE UPPER(name) = '" & UCase$(UserName) & "';"

        Set Database_RecordSet = Database_Connection.Execute(query)

        If Database_RecordSet.BOF Or Database_RecordSet.EOF Then
            Call WriteConsoleMsg(sendIndex, "Pj Inexistente", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        Call WriteConsoleMsg(sendIndex, "Pj: " & UserName, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "CiudadanosMatados: " & Database_RecordSet!ciudadanos_matados & ", CriminalesMatados: " & Database_RecordSet!criminales_matados & ", UsuariosMatados: " & Database_RecordSet!killed_users, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "NPCs muertos: " & Database_RecordSet!killed_npcs, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Clase: " & ListaClases(Database_RecordSet!class_id), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Genero: " & IIf(CByte(Database_RecordSet!ciudadanos_matados) = eGenero.Hombre, "Hombre", "Mujer"), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Raza: " & ListaRazas(Database_RecordSet!race_id), FontTypeNames.FONTTYPE_INFO)

        Set Database_RecordSet = Nothing
        
    #If DBConexionUnica = 0 Then
        Call Database_Close
    #End If

    End If

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in SendUserMiniStatsTxtFromDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub SendUserOROTxtFromDatabase(ByVal sendIndex As Integer, _
                                      ByVal UserName As String)

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 11/10/2018
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    If Not PersonajeExiste(UserName) Then
        Call WriteConsoleMsg(sendIndex, "Pj Inexistente", FontTypeNames.FONTTYPE_INFO)
    Else
        #If DBConexionUnica = 0 Then
            Call Database_Connect
        #Else
            'Si perdimos la conexion reconectamos
            If CheckSQLStatus = False Then Database_Reconnect
        #End If

        query = "SELECT bank_gold FROM usuario WHERE UPPER(name) = '" & UCase$(UserName) & "';"

        Set Database_RecordSet = Database_Connection.Execute(query)

        If Database_RecordSet.BOF Or Database_RecordSet.EOF Then
            Call WriteConsoleMsg(sendIndex, "Pj Inexistente", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        Call WriteConsoleMsg(sendIndex, "Pj: " & UserName, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Oro en banco: " & Database_RecordSet!bank_gold, FontTypeNames.FONTTYPE_INFO)

        Set Database_RecordSet = Nothing
        
        #If DBConexionUnica = 0 Then
            Call Database_Close
        #End If

    End If

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in SendUserOROTxtFromDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub SendUserInvTxtFromDatabase(ByVal sendIndex As Integer, _
                                      ByVal UserName As String)

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 11/10/2018
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query  As String

    Dim ObjInd As Long

    If Not PersonajeExiste(UserName) Then
        Call WriteConsoleMsg(sendIndex, "Pj Inexistente", FontTypeNames.FONTTYPE_INFO)
    Else
        #If DBConexionUnica = 0 Then
            Call Database_Connect
        #Else
            'Si perdimos la conexion reconectamos
            If CheckSQLStatus = False Then Database_Reconnect
        #End If

        query = "SELECT number, item_id, amount FROM inventory_item WHERE user_id = (SELECT id FROM usuario WHERE UPPER(name) = '" & UCase$(UserName) & "')"

        Set Database_RecordSet = Database_Connection.Execute(query)

        If Not Database_RecordSet.RecordCount = 0 Then
            Database_RecordSet.MoveFirst

            While Not Database_RecordSet.EOF

                ObjInd = val(Database_RecordSet!item_id)

                If ObjInd > 0 Then
                    Call WriteConsoleMsg(sendIndex, "Objeto " & Database_RecordSet!Number & " " & ObjData(ObjInd).Name & " Cantidad:" & Database_RecordSet!Amount, FontTypeNames.FONTTYPE_INFO)

                End If

                Database_RecordSet.MoveNext
            Wend
        Else
            Call WriteConsoleMsg(sendIndex, "Pj Inexistente", FontTypeNames.FONTTYPE_INFO)

        End If

        Set Database_RecordSet = Nothing
        
        #If DBConexionUnica = 0 Then
            Call Database_Close
        #End If

    End If

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in SendUserInvTxtFromDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub SendUserBovedaTxtFromDatabase(ByVal sendIndex As Integer, _
                                         ByVal UserName As String)

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 11/10/2018
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query  As String

    Dim ObjInd As Long

    If Not PersonajeExiste(UserName) Then
        Call WriteConsoleMsg(sendIndex, "Pj Inexistente", FontTypeNames.FONTTYPE_INFO)
    Else
        #If DBConexionUnica = 0 Then
            Call Database_Connect
        #Else
            'Si perdimos la conexion reconectamos
            If CheckSQLStatus = False Then Database_Reconnect
        #End If

        query = "SELECT number, item_id, amount FROM bank_item WHERE user_id = (SELECT id FROM usuario WHERE UPPER(name) = '" & UCase$(UserName) & "')"

        Set Database_RecordSet = Database_Connection.Execute(query)

        If Not Database_RecordSet.RecordCount = 0 Then
            Database_RecordSet.MoveFirst

            While Not Database_RecordSet.EOF

                ObjInd = val(Database_RecordSet!item_id)

                If ObjInd > 0 Then
                    Call WriteConsoleMsg(sendIndex, "Objeto " & Database_RecordSet!Number & " " & ObjData(ObjInd).Name & " Cantidad:" & Database_RecordSet!Amount, FontTypeNames.FONTTYPE_INFO)

                End If

                Database_RecordSet.MoveNext
            Wend
        Else
            Call WriteConsoleMsg(sendIndex, "Pj Inexistente", FontTypeNames.FONTTYPE_INFO)

        End If

        Set Database_RecordSet = Nothing
        
        #If DBConexionUnica = 0 Then
            Call Database_Close
        #End If

    End If

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in SendUserBovedaTxtFromDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub SendCharacterInfoDatabase(ByVal UserIndex As Integer, ByVal UserName As String)

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 11/10/2018
    '***************************************************
    On Error GoTo ErrorHandler

    Dim gName       As String

    Dim Miembro     As String

    Dim GuildActual As Integer

    Dim query       As String

    #If DBConexionUnica = 0 Then
        Call Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If CheckSQLStatus = False Then Database_Reconnect
    #End If

    query = "SELECT race_id, class_id, genre_id, level, gold, bank_gold, rep_average, guild_requests_history, guild_index, guild_member_history, pertenece_real, pertenece_caos, ciudadanos_matados, criminales_matados FROM usuario WHERE UPPER(name) = '" & UCase$(UserName) & "';"

    Set Database_RecordSet = Database_Connection.Execute(query)

    If Database_RecordSet.BOF Or Database_RecordSet.EOF Then
        Call WriteConsoleMsg(UserIndex, "Pj Inexistente", FontTypeNames.FONTTYPE_INFO)
        Exit Sub

    End If

    ' Get the character's current guild
    GuildActual = SanitizeNullValue(Database_RecordSet!Guild_Index, 0)

    If GuildActual > 0 And GuildActual <= CANTIDADDECLANES Then
        gName = "<" & GuildName(GuildActual) & ">"
    Else
        gName = "Ninguno"

    End If

    'Get previous guilds
    Miembro = SanitizeNullValue(Database_RecordSet!guild_member_history, vbNullString)

    If Len(Miembro) > 400 Then
        Miembro = ".." & Right$(Miembro, 400)

    End If

    Call Protocol.WriteCharacterInfo(UserIndex, UserName, Database_RecordSet!race_id, Database_RecordSet!class_id, Database_RecordSet!genre_id, Database_RecordSet!level, Database_RecordSet!Gold, Database_RecordSet!bank_gold, Database_RecordSet!rep_average, SanitizeNullValue(Database_RecordSet!guild_requests_history, vbNullString), gName, Miembro, Database_RecordSet!pertenece_real, Database_RecordSet!pertenece_caos, Database_RecordSet!ciudadanos_matados, Database_RecordSet!criminales_matados)

#If DBConexionUnica = 0 Then
    Call Database_Close
#End If

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in SendCharacterInfoDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Function GetUserGuildMemberDatabase(ByVal UserName As String) As String

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 11/10/2018
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    #If DBConexionUnica = 0 Then
        Call Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If CheckSQLStatus = False Then Database_Reconnect
    #End If

    query = "SELECT guild_member_history FROM usuario WHERE UPPER(name) = '" & UCase$(UserName) & "';"

    Set Database_RecordSet = Database_Connection.Execute(query)

    If Database_RecordSet.BOF Or Database_RecordSet.EOF Then
        GetUserGuildMemberDatabase = vbNullString
        Exit Function

    End If

    GetUserGuildMemberDatabase = SanitizeNullValue(Database_RecordSet!guild_member_history, vbNullString)
    Set Database_RecordSet = Nothing
    
    #If DBConexionUnica = 0 Then
        Call Database_Close
    #End If

    Exit Function

ErrorHandler:
    Call LogDatabaseError("Error in GetUserGuildMemberDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Function

Public Function GetUserGuildAspirantDatabase(ByVal UserName As String) As Integer

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 11/10/2018
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    #If DBConexionUnica = 0 Then
        Call Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If CheckSQLStatus = False Then Database_Reconnect
    #End If

    query = "SELECT guild_aspirant_index FROM usuario WHERE UPPER(name) = '" & UCase$(UserName) & "';"

    Set Database_RecordSet = Database_Connection.Execute(query)

    If Database_RecordSet.BOF Or Database_RecordSet.EOF Then
        GetUserGuildAspirantDatabase = 0
        Exit Function

    End If

    GetUserGuildAspirantDatabase = SanitizeNullValue(Database_RecordSet!guild_aspirant_index, 0)
    Set Database_RecordSet = Nothing
    
#If DBConexionUnica = 0 Then
    Call Database_Close
#End If

    Exit Function

ErrorHandler:
    Call LogDatabaseError("Error in GetUserGuildAspirantDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Function

Public Function GetUserGuildRejectionReasonDatabase(ByVal UserName As String) As String

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 11/10/2018
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    #If DBConexionUnica = 0 Then
        Call Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If CheckSQLStatus = False Then Database_Reconnect
    #End If

    query = "SELECT guild_rejected_because FROM usuario WHERE UPPER(name) = '" & UCase$(UserName) & "';"

    Set Database_RecordSet = Database_Connection.Execute(query)

    If Database_RecordSet.BOF Or Database_RecordSet.EOF Then
        GetUserGuildRejectionReasonDatabase = vbNullString
        Exit Function

    End If

    GetUserGuildRejectionReasonDatabase = SanitizeNullValue(Database_RecordSet!guild_rejected_because, vbNullString)
    Set Database_RecordSet = Nothing
    
#If DBConexionUnica = 0 Then
    Call Database_Close
#End If

    Exit Function

ErrorHandler:
    Call LogDatabaseError("Error in GetUserGuildRejectionReasonDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Function

Public Function GetUserGuildPedidosDatabase(ByVal UserName As String) As String

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 11/10/2018
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    #If DBConexionUnica = 0 Then
        Call Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If CheckSQLStatus = False Then Database_Reconnect
    #End If

    query = "SELECT guild_requests_history FROM usuario WHERE UPPER(name) = '" & UCase$(UserName) & "';"

    Set Database_RecordSet = Database_Connection.Execute(query)

    If Database_RecordSet.BOF Or Database_RecordSet.EOF Then
        GetUserGuildPedidosDatabase = vbNullString
        Exit Function

    End If

    GetUserGuildPedidosDatabase = SanitizeNullValue(Database_RecordSet!guild_requests_history, vbNullString)
    Set Database_RecordSet = Nothing
    
#If DBConexionUnica = 0 Then
    Call Database_Close
#End If

    Exit Function

ErrorHandler:
    Call LogDatabaseError("Error in GetUserGuildPedidosDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Function

Public Sub SaveUserGuildRejectionReasonDatabase(ByVal UserName As String, _
                                                ByVal Reason As String)

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 11/10/2018
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    #If DBConexionUnica = 0 Then
        Call Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If CheckSQLStatus = False Then Database_Reconnect
    #End If

    query = "UPDATE usuario SET "
    query = query & "guild_rejected_because = '" & Reason & "' "
    query = query & "WHERE UPPER(name) = '" & UCase$(UserName) & "';"

    Database_Connection.Execute (query)
    
#If DBConexionUnica = 0 Then
    Call Database_Close
#End If

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in SaveUserGuildRejectionReasonDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub SaveUserGuildIndexDatabase(ByVal UserName As String, _
                                      ByVal GuildIndex As Integer)

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 11/10/2018
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    #If DBConexionUnica = 0 Then
        Call Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If CheckSQLStatus = False Then Database_Reconnect
    #End If

    query = "UPDATE usuario SET "
    query = query & "guild_index = " & GuildIndex & " "
    query = query & "WHERE UPPER(name) = '" & UCase$(UserName) & "';"

    Database_Connection.Execute (query)
    
#If DBConexionUnica = 0 Then
    Call Database_Close
#End If

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in SaveUserGuildIndexDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub SaveUserGuildAspirantDatabase(ByVal UserName As String, _
                                         ByVal AspirantIndex As Integer)

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 11/10/2018
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    #If DBConexionUnica = 0 Then
        Call Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If CheckSQLStatus = False Then Database_Reconnect
    #End If

    query = "UPDATE usuario SET "
    query = query & "guild_aspirant_index = " & AspirantIndex & " "
    query = query & "WHERE UPPER(name) = '" & UCase$(UserName) & "';"

    Database_Connection.Execute (query)

#If DBConexionUnica = 0 Then
    Call Database_Close
#End If

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in SaveUserGuildAspirantDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub SaveUserGuildMemberDatabase(ByVal UserName As String, ByVal guilds As String)

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 11/10/2018
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    #If DBConexionUnica = 0 Then
        Call Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If CheckSQLStatus = False Then Database_Reconnect
    #End If

    query = "UPDATE usuario SET "
    query = query & "guild_member_history = '" & guilds & "' "
    query = query & "WHERE UPPER(name) = '" & UCase$(UserName) & "';"

    Database_Connection.Execute (query)

    #If DBConexionUnica = 0 Then
        Call Database_Close
    #End If

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in SaveUserGuildMemberDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub SaveUserGuildPedidosDatabase(ByVal UserName As String, ByVal Pedidos As String)

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 11/10/2018
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    #If DBConexionUnica = 0 Then
        Call Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If CheckSQLStatus = False Then Database_Reconnect
    #End If

    query = "UPDATE usuario SET "
    query = query & "guild_requests_history = '" & Pedidos & "' "
    query = query & "WHERE UPPER(name) = '" & UCase$(UserName) & "';"

    Database_Connection.Execute (query)

    #If DBConexionUnica = 0 Then
        Call Database_Close
    #End If

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in SaveUserGuildPedidosDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub SaveAccountLastLoginDatabase(ByVal UserName As String, ByVal UserIP As String)

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

    query = "UPDATE account SET "
    query = query & "date_last_login = NOW(), "
    query = query & "last_ip = '" & UserIP & "' "
    query = query & "WHERE UPPER(username) = '" & UCase$(UserName) & "';"

    Database_Connection.Execute (query)

    #If DBConexionUnica = 0 Then
        Call Database_Close
    #End If

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in SaveAccountLastLoginDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub SaveAccountEditGemasDatabase(ByVal UserName As String, ByVal Gemas As Long)

    '***************************************************
    'Author: Lorwik
    'Last Modification: 30/04/2020
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    #If DBConexionUnica = 0 Then
        Call Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If CheckSQLStatus = False Then Database_Reconnect
    #End If
    
    query = "UPDATE account SET gemas = '" & Gemas & "' WHERE id = (SELECT account_id FROM usuario WHERE UPPER(name) = '" & UCase$(UserName) & "');"

    Database_Connection.Execute (query)

    #If DBConexionUnica = 0 Then
        Call Database_Close
    #End If

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in SaveAccountLastLoginDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub SaveAccountSumaGemasDatabase(ByVal UserName As String, ByVal Gemas As Long)

    '***************************************************
    'Author: Lorwik
    'Last Modification: 30/04/2020
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String
    
    #If DBConexionUnica = 0 Then
        Call Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If CheckSQLStatus = False Then Database_Reconnect
    #End If

    query = "UPDATE account SET gemas = gemas + '" & Gemas & "' WHERE id = (SELECT account_id FROM usuario WHERE UPPER(name) = '" & UCase$(UserName) & "');"

    Database_Connection.Execute (query)

    #If DBConexionUnica = 0 Then
        Call Database_Close
    #End If

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in SaveAccountLastLoginDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub SaveAccountRestaGemasDatabase(ByVal UserName As String, ByVal Gemas As Long)

    '***************************************************
    'Author: Lorwik
    'Last Modification: 30/04/2020
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    #If DBConexionUnica = 0 Then
        Call Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If CheckSQLStatus = False Then Database_Reconnect
    #End If

    query = "UPDATE account SET gemas = gemas - '" & Gemas & "' WHERE id = (SELECT account_id FROM usuario WHERE UPPER(name) = '" & UCase$(UserName) & "');"

    Database_Connection.Execute (query)

    #If DBConexionUnica = 0 Then
        Call Database_Close
    #End If

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in SaveAccountLastLoginDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Function GetGemasDatabase(ByVal UserName As String) As Long

    '***************************************************
    'Author: Lorwik
    'Last Modification: 30/04/2020
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    #If DBConexionUnica = 0 Then
        Call Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If CheckSQLStatus = False Then Database_Reconnect
    #End If

    query = "SELECT gemas FROM account WHERE id = (SELECT account_id FROM usuario WHERE UPPER(name) = '" & UCase$(UserName) & "');"
    Debug.Print query
    Set Database_RecordSet = Database_Connection.Execute(query)

    If Database_RecordSet.BOF Or Database_RecordSet.EOF Then
        GetGemasDatabase = 0
        Exit Function

    End If

    GetGemasDatabase = CLng(Database_RecordSet!Gemas)
    Set Database_RecordSet = Nothing
    
    #If DBConexionUnica = 0 Then
        Call Database_Close
    #End If

    Exit Function

ErrorHandler:
    Call LogDatabaseError("Error in GetUserPromedioDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Function

Public Function SanitizeNullValue(ByVal Value As Variant, _
                                  ByVal defaultValue As Variant) As Variant
    SanitizeNullValue = IIf(IsNull(Value), defaultValue, Value)

End Function
