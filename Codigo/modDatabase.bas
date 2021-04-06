Attribute VB_Name = "modDatabase"
'Modulo gestor de la base de datos.
'Adaptado y mejorado por Lorwik

Option Explicit

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

    Dim LoopC  As Integer

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    'Basic user data
    With UserList(UserIndex)
        query = "INSERT INTO personaje SET "
        query = query & "name = '" & .Name & "', "
        query = query & "cuenta_id = " & .AccountInfo.ID & ", "
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
        Call User_Database.Database_Connection.Execute(query)

        'Get the user ID
        Set User_Database.Database_RecordSet = User_Database.Database_Connection.Execute("SELECT LAST_INSERT_ID();")

        If User_Database.Database_RecordSet.BOF Or User_Database.Database_RecordSet.EOF Then
            UserID = 1

        End If

        UserID = val(User_Database.Database_RecordSet.Fields(0).Value)
        Set User_Database.Database_RecordSet = Nothing

        .ID = UserID

        '*******************************************************************
        'Atributos
        '*******************************************************************
        query = "INSERT INTO atributos (user_id, "

        For LoopC = 1 To NUMATRIBUTOS
            query = query & " att" & LoopC
            If LoopC < NUMATRIBUTOS Then query = query & ", "
        Next LoopC

        query = query & ") VALUES (" & .ID & ", "

        For LoopC = 1 To NUMATRIBUTOS
            query = query & .Stats.UserAtributos(LoopC)
            If LoopC < NUMATRIBUTOS Then query = query & ", "

        Next LoopC
        
        query = query & ");"

        Call User_Database.Database_Connection.Execute(query)

        '*******************************************************************
        'Hechizos
        '*******************************************************************
        query = "INSERT INTO spell (user_id, "

        For LoopC = 1 To MAXUSERHECHIZOS
            query = query & " spell_id" & LoopC
            If LoopC < MAXUSERHECHIZOS Then query = query & ", "
        Next LoopC

        query = query & ") VALUES (" & .ID & ", "

        For LoopC = 1 To MAXUSERHECHIZOS
            query = query & .Stats.UserHechizos(LoopC)
            If LoopC < MAX_INVENTORY_SLOTS Then query = query & ", "

        Next LoopC
        
        query = query & ");"

        Call User_Database.Database_Connection.Execute(query)

        '*******************************************************************
        'Inventario
        '*******************************************************************
        query = "INSERT INTO inventario_items (user_id, "
        
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

        Call User_Database.Database_Connection.Execute(query)

        '*******************************************************************
        'Boveda
        '*******************************************************************
        query = "INSERT INTO banco_items (user_id, "
        
        For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
            query = query & "item_id" & LoopC & ", amount" & LoopC
            If LoopC < MAX_BANCOINVENTORY_SLOTS Then query = query & ", "
        Next LoopC
        
        query = query & ") VALUES (" & .ID & ", "
        
        For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS

            query = query & .BancoInvent.Object(LoopC).ObjIndex & ", "
            query = query & .BancoInvent.Object(LoopC).Amount
            If LoopC < MAX_BANCOINVENTORY_SLOTS Then query = query & ", "
            
        Next LoopC
        
        query = query & ");"

        Call User_Database.Database_Connection.Execute(query)

        '*******************************************************************
        'Skills
        '*******************************************************************
        query = "INSERT INTO skillpoint (user_id, "
        
        For LoopC = 1 To NUMSKILLS
            query = query & "sk" & LoopC & ", exp" & LoopC & ", elu" & LoopC
            If LoopC < NUMSKILLS Then query = query & ", "
        Next LoopC
        
        query = query & ") VALUES (" & .ID & ", "

        For LoopC = 1 To NUMSKILLS
            query = query & .Stats.UserSkills(LoopC) & ", "
            query = query & .Stats.ExpSkills(LoopC) & ", "
            query = query & .Stats.EluSkills(LoopC)
            If LoopC < NUMSKILLS Then query = query & ", "

        Next LoopC
        
        query = query & ");"

        Call User_Database.Database_Connection.Execute(query)
        
        '*******************************************************************
        'Quests
        '*******************************************************************
        query = "INSERT INTO quest (user_id, "
        
        For LoopC = 1 To MAXQUESTS
        
            query = query & "npcs" & LoopC & ", estado" & LoopC
            If LoopC < MAXQUESTS Then query = query & ", "
        
        Next LoopC
        
        query = query & ") VALUES (" & .ID & ", "
        
        For LoopC = 1 To MAXQUESTS
            query = query & "0, 0"
            If LoopC < MAXQUESTS Then query = query & ", "
            
        Next LoopC
        
        query = query & ");"
        
        Call User_Database.Database_Connection.Execute(query)
        
        '*******************************************************************
        'Profesion primaria
        '*******************************************************************
        query = "INSERT INTO profesion_primaria (user_id, profesion, "
        
        For LoopC = 1 To MAXUSERRECETAS
            query = query & "receta" & LoopC
            If LoopC < MAXUSERRECETAS Then query = query & ", "
        Next LoopC

        query = query & ") VALUES (" & .ID & ", " & .Profesion(0).Profesion & ", "

        For LoopC = 1 To MAXUSERRECETAS
            query = query & .Profesion(0).Recetas(LoopC)
            If LoopC < MAXUSERRECETAS Then query = query & ", "
        Next LoopC

        query = query & ");"

        Call User_Database.Database_Connection.Execute(query)
        
        '*******************************************************************
        'Profesion secundaria
        '*******************************************************************
        query = "INSERT INTO profesion_secundaria (user_id, profesion, "
        
        For LoopC = 1 To MAXUSERRECETAS
            query = query & "receta" & LoopC
            If LoopC < MAXUSERRECETAS Then query = query & ", "
        Next LoopC

        query = query & ") VALUES (" & .ID & ", " & .Profesion(1).Profesion & ", "

        For LoopC = 1 To MAXUSERRECETAS
            query = query & .Profesion(1).Recetas(LoopC)
            If LoopC < MAXUSERRECETAS Then query = query & ", "
        Next LoopC

        query = query & ");"

        Call User_Database.Database_Connection.Execute(query)
        
        '*******************************************************************
        'Mascotas
        '*******************************************************************
        query = "INSERT INTO pet (user_id, "
        
        For LoopC = 1 To MAXMASCOTAS
            query = query & "pet" & LoopC
            If LoopC < MAXMASCOTAS Then query = query & ", "
        Next LoopC

        query = query & ") VALUES (" & .ID & ", "

        For LoopC = 1 To MAXMASCOTAS
            query = query & .MascotasIndex(LoopC)
            If LoopC < MAXMASCOTAS Then query = query & ", "
        Next LoopC

        query = query & ");"

        Call User_Database.Database_Connection.Execute(query)
        
        '*******************************************************************
        'Amigos
        '*******************************************************************
        query = "INSERT INTO amigos (user_id) VALUES (" & .ID & ");"

        Call User_Database.Database_Connection.Execute(query)
        
    End With
    
    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
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
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    'Basic user data
    With UserList(UserIndex)
        query = "UPDATE personaje SET "
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
        Call User_Database.Database_Connection.Execute(query)

        '*******************************************************************
        'Hechizos
        '*******************************************************************
        
        query = "UPDATE spell SET "
        
        For LoopC = 1 To MAXUSERHECHIZOS
            
            query = query & "spell_id" & LoopC & " = '" & .Stats.UserHechizos(LoopC) & "' "
            If LoopC < MAXUSERHECHIZOS Then query = query & ", "
            
        Next LoopC
        
        query = query & " WHERE user_id = '" & .ID & "'"

        Call User_Database.Database_Connection.Execute(query)

        '*******************************************************************
        'Inventario
        '*******************************************************************
        
        query = "UPDATE inventario_items SET "
        
        For LoopC = 1 To MAX_INVENTORY_SLOTS
            
            query = query & "item_id" & LoopC & " = '" & .Invent.Object(LoopC).ObjIndex & "', "
            query = query & "amount" & LoopC & " = '" & .Invent.Object(LoopC).Amount & "', "
            query = query & "is_equipped" & LoopC & " = '" & .Invent.Object(LoopC).Equipped & "'"
            
            If LoopC < MAX_INVENTORY_SLOTS Then query = query & ", "
            
        Next LoopC
        
        query = query & " WHERE user_id = '" & .ID & "'"
        
        Call User_Database.Database_Connection.Execute(query)

        '*******************************************************************
        'Boveda
        '*******************************************************************
        
        query = "UPDATE banco_items SET "
        
        For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
            
            query = query & "item_id" & LoopC & " = '" & .BancoInvent.Object(LoopC).ObjIndex & "', "
            query = query & "amount" & LoopC & " = '" & .BancoInvent.Object(LoopC).Amount & "'"
            
            If LoopC < MAX_BANCOINVENTORY_SLOTS Then query = query & ", "
            
        Next LoopC
        
        query = query & " WHERE user_id = '" & .ID & "'"
        
        Call User_Database.Database_Connection.Execute(query)

        '*******************************************************************
        'Skills
        '*******************************************************************
        query = "UPDATE skillpoint SET "
        
        For LoopC = 1 To NUMSKILLS
            
            query = query & "sk" & LoopC & " = '" & .Stats.UserSkills(LoopC) & "', "
            query = query & "exp" & LoopC & " = '" & .Stats.ExpSkills(LoopC) & "', "
            query = query & "elu" & LoopC & " = '" & .Stats.EluSkills(LoopC) & "'"
            If LoopC < NUMSKILLS Then query = query & ", "

        Next LoopC
        
        query = query & " WHERE user_id = '" & .ID & "'"
        
        Call User_Database.Database_Connection.Execute(query)
       
        '*******************************************************************
        'Profesion primaria
        '*******************************************************************
        query = "UPDATE profesion_primaria SET profesion" & " = '" & .Profesion(0).Profesion & "', "
        
        For LoopC = 1 To MAXUSERRECETAS
            query = query & "receta" & LoopC & " = '" & .Profesion(0).Recetas(LoopC) & "'"
            If LoopC < MAXUSERRECETAS Then query = query & ", "
        Next LoopC
        
        query = query & " WHERE user_id = '" & .ID & "'"
            
        Call User_Database.Database_Connection.Execute(query)
        
        '*******************************************************************
        'Profesion secundaria
        '*******************************************************************
        query = "UPDATE profesion_secundaria SET profesion" & " = '" & .Profesion(1).Profesion & "', "
        
        For LoopC = 1 To MAXUSERRECETAS
            query = query & "receta" & LoopC & " = '" & .Profesion(1).Recetas(LoopC) & "'"
            If LoopC < MAXUSERRECETAS Then query = query & ", "
        Next LoopC
        
        query = query & " WHERE user_id = '" & .ID & "'"
            
        Call User_Database.Database_Connection.Execute(query)
        
        '*******************************************************************
        'Mascotas
        '*******************************************************************
        Dim petType As Integer
        
        query = "UPDATE pet SET "
        
        For LoopC = 1 To MAXMASCOTAS
            
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

            query = query & "pet" & LoopC & " = '" & petType & "' "
            If LoopC < MAXMASCOTAS Then query = query & ", "
        Next LoopC
        
        query = query & "WHERE user_id = '" & .ID & "'"

        Call User_Database.Database_Connection.Execute(query)
        
        '*******************************************************************
        'Amigos
        '*******************************************************************
        query = "UPDATE amigos SET "
        
        For LoopC = 1 To MAXAMIGOS
            query = query & "amigo" & LoopC & " = '" & .Amigos(LoopC).Nombre & "', "
            query = query & "ignorado" & LoopC & " = '" & .Amigos(LoopC).Ignorado & "'"
            If LoopC < MAXAMIGOS Then query = query & ", "
        Next LoopC
        
        query = query & " WHERE user_id = '" & .ID & "'"
            
        Call User_Database.Database_Connection.Execute(query)

    End With

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
    #End If

    Exit Sub

ErrorHandler:
    Call LogDatabaseError("Unable to UPDATE personaje to Mysql Database: " & UserList(UserIndex).Name & ". " & Err.Number & " - " & Err.description)

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
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If
    
    'Basic user data
    With UserList(UserIndex)

        query = "UPDATE quest SET "

        For LoopC = 1 To MAXQUESTS
            
            query = query & "estado" & LoopC & " = '" & CInt(.QuestStats.Quests(LoopC).QuestStatus) & "', "           'Estado de la quest
            
            tmpst = "npcs" & LoopC & " = '0' "
            
            If .QuestStats.Quests(LoopC).QuestStatus = eStatusQuest.EnCurso Then
            
                '¿La quest requiere matar NPC?
                If QuestList(LoopC).RequiredNPCs > 0 Then
                    tmpst = "npcs" & LoopC & " = '"
                    For j = 1 To QuestList(LoopC).RequiredNPCs
                    
                        tmpst = tmpst & val(.QuestStats.Quests(LoopC).NPCsKilled(j))   'Cuantos NPCs se ha matado de los que requeridos
                        
                        If Not j = QuestList(LoopC).RequiredNPCs Then tmpst = tmpst & "."
                    Next j
                    
                    tmpst = tmpst & "' "

                End If
                
            End If
            
            If LoopC < MAXQUESTS Then tmpst = tmpst & ","
            
            query = query & tmpst

        Next LoopC
        
        query = query & " WHERE user_id = '" & .ID & "'"

        Call User_Database.Database_Connection.Execute(query)
        
        'Debug.Print query

    End With

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
    #End If

    Exit Sub

ErrorHandler:
    Call LogDatabaseError("Unable to UPDATE personaje to Mysql Database: " & UserList(UserIndex).Name & ". " & Err.Number & " - " & Err.description)
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
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    'Basic user data
    With UserList(UserIndex)
        query = "SELECT *, DATE_FORMAT(fecha_ingreso, '%Y-%m-%d') as 'fecha_ingreso_format' FROM personaje WHERE UPPER(name) ='" & UCase$(.Name) & "';"
        Set User_Database.Database_RecordSet = User_Database.Database_Connection.Execute(query)

        If User_Database.Database_RecordSet.BOF Or User_Database.Database_RecordSet.EOF Then Exit Sub

        'Start setting data
        .ID = User_Database.Database_RecordSet!ID
        .Name = User_Database.Database_RecordSet!Name
        .Stats.ELV = User_Database.Database_RecordSet!level
        .Stats.Exp = User_Database.Database_RecordSet!Exp
        .Stats.ELU = User_Database.Database_RecordSet!ELU
        .Genero = User_Database.Database_RecordSet!genre_id
        .Raza = User_Database.Database_RecordSet!race_id
        .clase = User_Database.Database_RecordSet!class_id
        .Hogar = User_Database.Database_RecordSet!home_id
        .Desc = User_Database.Database_RecordSet!description
        .Stats.Gld = User_Database.Database_RecordSet!Gold
        .Stats.Banco = User_Database.Database_RecordSet!bank_gold
        .Stats.SkillPts = User_Database.Database_RecordSet!free_skillpoints
        .Counters.AsignedSkills = User_Database.Database_RecordSet!assigned_skillpoints
        .Stats.ELO = User_Database.Database_RecordSet!ELO
        .NroMascotas = User_Database.Database_RecordSet!pet_amount
        .Pos.Map = User_Database.Database_RecordSet!pos_map
        .Pos.X = User_Database.Database_RecordSet!pos_x
        .Pos.Y = User_Database.Database_RecordSet!pos_y
        .flags.lastMap = User_Database.Database_RecordSet!last_map
        .OrigChar.body = User_Database.Database_RecordSet!body_id
        .OrigChar.Head = User_Database.Database_RecordSet!head_id
        .OrigChar.WeaponAnim = User_Database.Database_RecordSet!weapon_id
        .OrigChar.CascoAnim = User_Database.Database_RecordSet!helmet_id
        .OrigChar.ShieldAnim = User_Database.Database_RecordSet!shield_id
        .OrigChar.Heading = User_Database.Database_RecordSet!Heading
        .OrigChar.AuraAnim = User_Database.Database_RecordSet!Aura_id
        .OrigChar.AuraColor = User_Database.Database_RecordSet!Aura_color
        .Invent.NroItems = User_Database.Database_RecordSet!items_amount
        .Invent.ArmourEqpSlot = SanitizeNullValue(User_Database.Database_RecordSet!slot_armour, 0)
        .Invent.WeaponEqpSlot = SanitizeNullValue(User_Database.Database_RecordSet!slot_weapon, 0)
        .Invent.CascoEqpSlot = SanitizeNullValue(User_Database.Database_RecordSet!slot_helmet, 0)
        .Invent.EscudoEqpSlot = SanitizeNullValue(User_Database.Database_RecordSet!slot_shield, 0)
        .Invent.MunicionEqpSlot = SanitizeNullValue(User_Database.Database_RecordSet!slot_ammo, 0)
        .Invent.BarcoSlot = SanitizeNullValue(User_Database.Database_RecordSet!slot_ship, 0)
        .Invent.AnilloEqpSlot = SanitizeNullValue(User_Database.Database_RecordSet!slot_ring, 0)
        .Invent.MochilaEqpSlot = SanitizeNullValue(User_Database.Database_RecordSet!slot_bag, 0)
        .Stats.MinHp = User_Database.Database_RecordSet!min_hp
        .Stats.MaxHp = User_Database.Database_RecordSet!max_hp
        .Stats.MinMAN = User_Database.Database_RecordSet!min_man
        .Stats.MaxMAN = User_Database.Database_RecordSet!max_man
        .Stats.MinSta = User_Database.Database_RecordSet!min_sta
        .Stats.MaxSta = User_Database.Database_RecordSet!max_sta
        .Stats.MinHam = User_Database.Database_RecordSet!min_ham
        .Stats.MaxHam = User_Database.Database_RecordSet!max_ham
        .Stats.MinAGU = User_Database.Database_RecordSet!min_sed
        .Stats.MaxAGU = User_Database.Database_RecordSet!max_sed
        .Stats.MinHIT = User_Database.Database_RecordSet!min_hit
        .Stats.MaxHIT = User_Database.Database_RecordSet!max_hit
        .Stats.NPCsMuertos = User_Database.Database_RecordSet!killed_npcs
        .Stats.UsuariosMatados = User_Database.Database_RecordSet!killed_users
        .Reputacion.AsesinoRep = User_Database.Database_RecordSet!rep_asesino
        .Reputacion.BandidoRep = User_Database.Database_RecordSet!rep_bandido
        .Reputacion.BurguesRep = User_Database.Database_RecordSet!rep_burgues
        .Reputacion.LadronesRep = User_Database.Database_RecordSet!rep_ladron
        .Reputacion.NobleRep = User_Database.Database_RecordSet!rep_noble
        .Reputacion.PlebeRep = User_Database.Database_RecordSet!rep_plebe
        .Reputacion.Promedio = User_Database.Database_RecordSet!rep_average
        .flags.Desnudo = User_Database.Database_RecordSet!is_naked
        .flags.Envenenado = User_Database.Database_RecordSet!is_poisoned
        .flags.Incinerado = User_Database.Database_RecordSet!is_incinerado
        .flags.Escondido = User_Database.Database_RecordSet!is_hidden
        .flags.Hambre = User_Database.Database_RecordSet!is_hungry
        .flags.Sed = User_Database.Database_RecordSet!is_thirsty
        .flags.Ban = User_Database.Database_RecordSet!is_ban
        .flags.Muerto = User_Database.Database_RecordSet!is_dead
        .flags.Navegando = User_Database.Database_RecordSet!is_sailing
        .flags.Paralizado = User_Database.Database_RecordSet!is_paralyzed
        .Counters.Pena = User_Database.Database_RecordSet!counter_pena
        .flags.Global = User_Database.Database_RecordSet!is_global
        .Profesion(0).Profesion = User_Database.Database_RecordSet!ProfesionA
        .Profesion(1).Profesion = User_Database.Database_RecordSet!ProfesionB
        .flags.ModoCombate = User_Database.Database_RecordSet!ModoCombate
        .flags.Seguro = User_Database.Database_RecordSet!Seguro
        
        If User_Database.Database_RecordSet!pertenece_consejo_real Then
            .flags.Privilegios = .flags.Privilegios Or PlayerType.RoyalCouncil

        End If

        If User_Database.Database_RecordSet!pertenece_consejo_caos Then
            .flags.Privilegios = .flags.Privilegios Or PlayerType.ChaosCouncil

        End If

        .Faccion.ArmadaReal = User_Database.Database_RecordSet!pertenece_real
        .Faccion.FuerzasCaos = User_Database.Database_RecordSet!pertenece_caos
        .Faccion.CiudadanosMatados = User_Database.Database_RecordSet!ciudadanos_matados
        .Faccion.CriminalesMatados = User_Database.Database_RecordSet!criminales_matados
        .Faccion.RecibioArmaduraReal = User_Database.Database_RecordSet!recibio_armadura_real
        .Faccion.RecibioArmaduraCaos = User_Database.Database_RecordSet!recibio_armadura_caos
        .Faccion.RecibioExpInicialReal = User_Database.Database_RecordSet!recibio_exp_real
        .Faccion.RecibioExpInicialCaos = User_Database.Database_RecordSet!recibio_exp_caos
        .Faccion.RecompensasReal = User_Database.Database_RecordSet!recompensas_real
        .Faccion.RecompensasCaos = User_Database.Database_RecordSet!recompensas_caos
        .Faccion.Reenlistadas = User_Database.Database_RecordSet!Reenlistadas
        .Faccion.FechaIngreso = SanitizeNullValue(User_Database.Database_RecordSet!fecha_ingreso_format, vbNullString)
        .Faccion.NivelIngreso = SanitizeNullValue(User_Database.Database_RecordSet!nivel_ingreso, 0)
        .Faccion.MatadosIngreso = SanitizeNullValue(User_Database.Database_RecordSet!matados_ingreso, 0)
        .Faccion.NextRecompensa = SanitizeNullValue(User_Database.Database_RecordSet!siguiente_recompensa, 0)

        .GuildIndex = SanitizeNullValue(User_Database.Database_RecordSet!Guild_Index, 0)

        Set User_Database.Database_RecordSet = Nothing

        '*******************************************************************
        'Atributos
        '*******************************************************************
        query = "SELECT * FROM atributos WHERE user_id = " & .ID & ";"
        Set User_Database.Database_RecordSet = User_Database.Database_Connection.Execute(query)
    
        If Not User_Database.Database_RecordSet.RecordCount = 0 Then
            
            User_Database.Database_RecordSet.MoveFirst
            
            For LoopC = 1 To NUMATRIBUTOS

                .Stats.UserAtributos(LoopC) = User_Database.Database_RecordSet("att" & LoopC)
                .Stats.UserAtributosBackUP(LoopC) = .Stats.UserAtributos(LoopC)

            Next LoopC

        End If

        Set User_Database.Database_RecordSet = Nothing

        '*******************************************************************
        'Hechizos
        '*******************************************************************
        query = "SELECT * FROM spell WHERE user_id = " & .ID & ";"
        Set User_Database.Database_RecordSet = User_Database.Database_Connection.Execute(query)

        If Not User_Database.Database_RecordSet.RecordCount = 0 Then
            User_Database.Database_RecordSet.MoveFirst

            For LoopC = 1 To MAXUSERHECHIZOS
                .Stats.UserHechizos(LoopC) = User_Database.Database_RecordSet("spell_id" & LoopC)
            Next LoopC

        End If

        Set User_Database.Database_RecordSet = Nothing

        '*******************************************************************
        'Mascotas
        '*******************************************************************
        query = "SELECT * FROM pet WHERE user_id = " & .ID & ";"
        Set User_Database.Database_RecordSet = User_Database.Database_Connection.Execute(query)

        If Not User_Database.Database_RecordSet.RecordCount = 0 Then
            User_Database.Database_RecordSet.MoveFirst

            For LoopC = 1 To MAXMASCOTAS
                .MascotasType(LoopC) = User_Database.Database_RecordSet("pet" & LoopC)
            Next LoopC

        End If

        Set User_Database.Database_RecordSet = Nothing

        '*******************************************************************
        'Inventario
        '*******************************************************************
        query = "SELECT * FROM inventario_items WHERE user_id = " & .ID & ";"
        Set User_Database.Database_RecordSet = User_Database.Database_Connection.Execute(query)

        If Not User_Database.Database_RecordSet.RecordCount = 0 Then
            User_Database.Database_RecordSet.MoveFirst
                
            For LoopC = 1 To MAX_INVENTORY_SLOTS
                .Invent.Object(LoopC).ObjIndex = User_Database.Database_RecordSet("item_id" & LoopC)
                .Invent.Object(LoopC).Amount = User_Database.Database_RecordSet("Amount" & LoopC)
                .Invent.Object(LoopC).Equipped = User_Database.Database_RecordSet("is_equipped" & LoopC)
            Next LoopC
                
        End If

        Set User_Database.Database_RecordSet = Nothing

        '*******************************************************************
        'Boveda
        '*******************************************************************
        query = "SELECT * FROM banco_items WHERE user_id = " & .ID & ";"
        Set User_Database.Database_RecordSet = User_Database.Database_Connection.Execute(query)

        If Not User_Database.Database_RecordSet.RecordCount = 0 Then
            User_Database.Database_RecordSet.MoveFirst
                
            For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
                .BancoInvent.Object(LoopC).ObjIndex = User_Database.Database_RecordSet("item_id" & LoopC)
                .BancoInvent.Object(LoopC).Amount = User_Database.Database_RecordSet("Amount" & LoopC)
            Next LoopC
                
        End If

        Set User_Database.Database_RecordSet = Nothing

        '*******************************************************************
        'Skills
        '*******************************************************************
        query = "SELECT * FROM skillpoint WHERE user_id = " & .ID & ";"
        Set User_Database.Database_RecordSet = User_Database.Database_Connection.Execute(query)

        If Not User_Database.Database_RecordSet.RecordCount = 0 Then
            User_Database.Database_RecordSet.MoveFirst

            For LoopC = 1 To NUMSKILLS
                .Stats.UserSkills(LoopC) = User_Database.Database_RecordSet("sk" & LoopC)
                .Stats.ExpSkills(LoopC) = User_Database.Database_RecordSet("exp" & LoopC)
                .Stats.EluSkills(LoopC) = User_Database.Database_RecordSet("elu" & LoopC)
            Next LoopC

        End If

        Set User_Database.Database_RecordSet = Nothing
        
        '*******************************************************************
        'Profesion primaria
        '*******************************************************************
        query = "SELECT * FROM profesion_primaria WHERE user_id = " & .ID & ";"
        Set User_Database.Database_RecordSet = User_Database.Database_Connection.Execute(query)

        If Not User_Database.Database_RecordSet.RecordCount = 0 Then
            User_Database.Database_RecordSet.MoveFirst

            .Profesion(0).Profesion = User_Database.Database_RecordSet("profesion")

            For LoopC = 1 To MAXUSERRECETAS
                .Profesion(0).Recetas(LoopC) = User_Database.Database_RecordSet("receta" & LoopC)
            Next LoopC

        End If

        Set User_Database.Database_RecordSet = Nothing
        
        '*******************************************************************
        'Profesion secundaria
        '*******************************************************************
        query = "SELECT * FROM profesion_secundaria WHERE user_id = " & .ID & ";"
        Set User_Database.Database_RecordSet = User_Database.Database_Connection.Execute(query)

        If Not User_Database.Database_RecordSet.RecordCount = 0 Then
            User_Database.Database_RecordSet.MoveFirst

            .Profesion(1).Profesion = User_Database.Database_RecordSet("profesion")

            For LoopC = 1 To MAXUSERRECETAS
                .Profesion(1).Recetas(LoopC) = User_Database.Database_RecordSet("receta" & LoopC)
            Next LoopC

        End If

        Set User_Database.Database_RecordSet = Nothing
        
        '*******************************************************************
        'Amigos
        '*******************************************************************
        query = "SELECT * FROM amigos WHERE user_id = " & .ID & ";"
        Set User_Database.Database_RecordSet = User_Database.Database_Connection.Execute(query)

        If Not User_Database.Database_RecordSet.RecordCount = 0 Then
            User_Database.Database_RecordSet.MoveFirst

            For LoopC = 1 To MAXAMIGOS
                .Amigos(LoopC).Nombre = User_Database.Database_RecordSet("amigo" & LoopC)
                .Amigos(LoopC).Ignorado = User_Database.Database_RecordSet("ignorado" & LoopC)
            Next LoopC

        End If

        Set User_Database.Database_RecordSet = Nothing

    End With

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
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

    Dim query           As String
    Dim Fields()        As String
    Dim j               As Integer
    Dim tmpint          As Integer
    Dim LoopC           As Integer
    Dim NPCRequeridos   As String
    
    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    With UserList(UserIndex).QuestStats

        query = "SELECT * FROM quest WHERE user_id = '" & UserList(UserIndex).ID & "';"
        Set User_Database.Database_RecordSet = User_Database.Database_Connection.Execute(query)
    
        If Not User_Database.Database_RecordSet.RecordCount = 0 Then
            User_Database.Database_RecordSet.MoveFirst
            
            For LoopC = 1 To NumQuests

                    '¿La quest requiere matar NPC?
                    If QuestList(LoopC).RequiredNPCs Then
                        ReDim .Quests(LoopC).NPCsKilled(1 To QuestList(LoopC).RequiredNPCs)
            
                        NPCRequeridos = User_Database.Database_RecordSet("npcs" & LoopC)
            
                        Fields = Split(NPCRequeridos, ".")
                        
                        For j = 1 To QuestList(LoopC).RequiredNPCs

                            If UBound(Fields()) > 0 Then
                                .Quests(LoopC).NPCsKilled(j) = CInt(Fields(j - 1))
                            Else
                                .Quests(LoopC).NPCsKilled(j) = NPCRequeridos
                            End If
                        Next j
     
                    .Quests(LoopC).QuestStatus = CByte(User_Database.Database_RecordSet("estado" & LoopC))
                             
                    'Si la quest actual se termino, lo sumamos al contador de terminados
                    If .Quests(LoopC).QuestStatus = eStatusQuest.Terminada Then .NumQuestsDone = .NumQuestsDone + 1
                        
                End If
                
            Next LoopC
                
           Call ListarQuestsenCurso(UserIndex)
                
        End If

    End With
    
    Set User_Database.Database_RecordSet = Nothing

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
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
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "SELECT id FROM personaje WHERE UPPER(name) = '" & UCase$(UserName) & "' AND deleted = FALSE;"

    Set User_Database.Database_RecordSet = User_Database.Database_Connection.Execute(query)

    If User_Database.Database_RecordSet.BOF Or User_Database.Database_RecordSet.EOF Then
        PersonajeExisteDatabase = False
        Exit Function

    End If

    PersonajeExisteDatabase = (User_Database.Database_RecordSet.RecordCount > 0)
    Set User_Database.Database_RecordSet = Nothing
    
    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
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
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "SELECT is_ban FROM personaje WHERE UPPER(name) = '" & UCase$(UserName) & "';"

    Set User_Database.Database_RecordSet = User_Database.Database_Connection.Execute(query)

    If User_Database.Database_RecordSet.BOF Or User_Database.Database_RecordSet.EOF Then
        BANCheckDatabase = False
        Exit Function

    End If

    BANCheckDatabase = CBool(User_Database.Database_RecordSet!is_ban)

    Set User_Database.Database_RecordSet = Nothing
    
    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
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
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "UPDATE personaje SET is_ban = FALSE WHERE UPPER(name) = '" & UCase$(UserName) & "';"

    User_Database.Database_Connection.Execute (query)

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
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
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "SELECT guild_index FROM personaje WHERE UPPER(name) = '" & UCase$(UserName) & "';"

    Set User_Database.Database_RecordSet = User_Database.Database_Connection.Execute(query)

    If User_Database.Database_RecordSet.BOF Or User_Database.Database_RecordSet.EOF Then
        GetUserGuildIndexDatabase = 0
        Exit Function

    End If

    GetUserGuildIndexDatabase = SanitizeNullValue(User_Database.Database_RecordSet!Guild_Index, 0)
    Set User_Database.Database_RecordSet = Nothing
    
    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
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
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "UPDATE personaje SET name = '" & UCase$(newName) & "' WHERE UPPER(name) = '" & UCase$(UserName) & "';"

    User_Database.Database_Connection.Execute (query)

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
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
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "UPDATE personaje SET votes_amount = " & NumeroEncuesta & " WHERE id = " & UserList(UserIndex).ID & ";"

    User_Database.Database_Connection.Execute (query)
    
    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
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
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "SELECT votes_amount FROM personaje WHERE UPPER(name) = '" & UCase$(UserName) & "';"

    Set User_Database.Database_RecordSet = User_Database.Database_Connection.Execute(query)

    If User_Database.Database_RecordSet.BOF Or User_Database.Database_RecordSet.EOF Then
        PersonajeCantidadVotosDatabase = 0
        Exit Function

    End If

    PersonajeCantidadVotosDatabase = CInt(User_Database.Database_RecordSet!votes_amount)
    Set User_Database.Database_RecordSet = Nothing
    
    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
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
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "UPDATE personaje SET is_ban = TRUE WHERE UPPER(name) = '" & UCase$(UserName) & "';"

    User_Database.Database_Connection.Execute (query)

    query = "INSERT INTO punishment SET "
    query = query & "user_id = (SELECT id FROM personaje WHERE UPPER(name) = '" & UCase$(UserName) & "'), "
    query = query & "number = " & (cantPenas + 1) & ", "
    query = query & "reason = '" & BannedBy & ": BAN POR " & LCase$(Reason) & " " & Date & " " & time & "';"

    User_Database.Database_Connection.Execute (query)

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
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
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "SELECT COUNT(1) as punishments FROM punishment WHERE user_id = (SELECT id FROM personaje WHERE UPPER(name) = '" & UCase$(UserName) & "')"

    Set User_Database.Database_RecordSet = User_Database.Database_Connection.Execute(query)

    If User_Database.Database_RecordSet.BOF Or User_Database.Database_RecordSet.EOF Then
        GetUserAmountOfPunishments = 0
        Exit Function

    End If

    GetUserAmountOfPunishments = CInt(User_Database.Database_RecordSet!punishments)
    Set User_Database.Database_RecordSet = Nothing
    
    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
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
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "SELECT * FROM punishment WHERE user_id = (SELECT id FROM personaje WHERE UPPER(name) = '" & UCase$(UserName) & "');"

    Set User_Database.Database_RecordSet = User_Database.Database_Connection.Execute(query)

    If Not User_Database.Database_RecordSet.RecordCount = 0 Then
        User_Database.Database_RecordSet.MoveFirst

        While Not User_Database.Database_RecordSet.EOF

            Call WriteConsoleMsg(UserIndex, User_Database.Database_RecordSet!Number & " - " & User_Database.Database_RecordSet!Reason, FontTypeNames.FONTTYPE_INFO)

            User_Database.Database_RecordSet.MoveNext
        Wend

    End If

    Set User_Database.Database_RecordSet = Nothing
    
    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
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
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "SELECT pos_map, pos_x, pos_y FROM personaje WHERE UPPER(name) = '" & UCase$(UserName) & "';"

    Set User_Database.Database_RecordSet = User_Database.Database_Connection.Execute(query)

    If User_Database.Database_RecordSet.BOF Or User_Database.Database_RecordSet.EOF Then
        GetUserPos = vbNullString
        Exit Function

    End If

    GetUserPos = User_Database.Database_RecordSet!pos_map & "-" & User_Database.Database_RecordSet!pos_x & "-" & User_Database.Database_RecordSet!pos_y
    Set User_Database.Database_RecordSet = Nothing
    
    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
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
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "INSERT INTO punishment SET "
    query = query & "user_id = (SELECT id FROM personaje WHERE UPPER(name) = '" & UCase$(UserName) & "'), "
    query = query & "number = " & Number & ", "
    query = query & "reason = '" & Reason & "';"

    User_Database.Database_Connection.Execute (query)

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
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
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "UPDATE punishment SET "
    query = query & "reason = '" & Reason & "' "
    query = query & "WHERE number = " & Number & " AND user_id = (SELECT id FROM personaje WHERE UPPER(name) = '" & UCase$(UserName) & "');"

    User_Database.Database_Connection.Execute (query)

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
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
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "UPDATE personaje SET "
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

    User_Database.Database_Connection.Execute (query)

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
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
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "UPDATE personaje SET "
    query = query & "pertenece_consejo_real = FALSE, "
    query = query & "pertenece_consejo_caos = FALSE "
    query = query & "WHERE UPPER(name) = '" & UCase$(UserName) & "';"

    User_Database.Database_Connection.Execute (query)

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
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
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "UPDATE personaje SET "
    query = query & "pertenece_real = FALSE, "
    query = query & "pertenece_caos = FALSE "
    query = query & "WHERE UPPER(name) = '" & UCase$(UserName) & "';"

    User_Database.Database_Connection.Execute (query)
    
    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
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
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "UPDATE personaje SET "
    query = query & "pertenece_caos = FALSE, "
    query = query & "reenlistadas = 200 "
    query = query & "WHERE UPPER(name) = '" & UCase$(UserName) & "';"

    User_Database.Database_Connection.Execute (query)

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
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
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "UPDATE personaje SET "
    query = query & "pertenece_real = FALSE, "
    query = query & "reenlistadas = 200 "
    query = query & "WHERE UPPER(name) = '" & UCase$(UserName) & "';"

    User_Database.Database_Connection.Execute (query)

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
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
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "UPDATE personaje SET "
    query = query & "is_logged = " & IIf(Logged = 1, "TRUE", "FALSE") & " "
    query = query & "WHERE UPPER(name) = '" & UCase$(UserName) & "';"

    User_Database.Database_Connection.Execute (query)

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
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
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "SELECT last_ip FROM cuentas WHERE id = (SELECT cuenta_id FROM personaje WHERE UPPER(name) = '" & UCase$(UserName) & "');"

    Set User_Database.Database_RecordSet = User_Database.Database_Connection.Execute(query)

    If User_Database.Database_RecordSet.BOF Or User_Database.Database_RecordSet.EOF Then
        GetUserLastIps = vbNullString
        Exit Function

    End If

    GetUserLastIps = User_Database.Database_RecordSet!last_ip
    Set User_Database.Database_RecordSet = Nothing
    
    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
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
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "SELECT number, value FROM skillpoint WHERE user_id = (SELECT id FROM personaje WHERE UPPER(name) = '" & UCase$(UserName) & "');"

    Set User_Database.Database_RecordSet = User_Database.Database_Connection.Execute(query)

    If Not User_Database.Database_RecordSet.RecordCount = 0 Then
        User_Database.Database_RecordSet.MoveFirst

        While Not User_Database.Database_RecordSet.EOF

            GetUserSkills = GetUserSkills & "CHAR>" & SkillsNames(User_Database.Database_RecordSet!Number) & " = " & User_Database.Database_RecordSet!Value & vbCrLf

            User_Database.Database_RecordSet.MoveNext
        Wend

    End If

    Set User_Database.Database_RecordSet = Nothing

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
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
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "SELECT free_skillpoints FROM personaje WHERE UPPER(name) = '" & UCase$(UserName) & "';"

    Set User_Database.Database_RecordSet = User_Database.Database_Connection.Execute(query)

    If User_Database.Database_RecordSet.BOF Or User_Database.Database_RecordSet.EOF Then
        GetUserFreeSkills = 0
        Exit Function

    End If

    GetUserFreeSkills = CInt(User_Database.Database_RecordSet!free_skillpoints)
    Set User_Database.Database_RecordSet = Nothing
    
    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
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
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "UPDATE personaje SET "
    query = query & "counter_training = " & trainingTime & " "
    query = query & "WHERE UPPER(name) = '" & UCase$(UserName) & "';"

    User_Database.Database_Connection.Execute (query)

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
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
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "SELECT counter_training FROM personaje WHERE UPPER(name) = '" & UCase$(UserName) & "';"

    Set User_Database.Database_RecordSet = User_Database.Database_Connection.Execute(query)

    If User_Database.Database_RecordSet.BOF Or User_Database.Database_RecordSet.EOF Then
        GetUserTrainingTime = 0
        Exit Function

    End If

    GetUserTrainingTime = CLng(User_Database.Database_RecordSet!counter_training)
    Set User_Database.Database_RecordSet = Nothing
    
    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
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
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "SELECT pertenece_real FROM personaje WHERE UPPER(name) = '" & UCase$(UserName) & "' AND deleted = FALSE;"

    Set User_Database.Database_RecordSet = User_Database.Database_Connection.Execute(query)

    If User_Database.Database_RecordSet.BOF Or User_Database.Database_RecordSet.EOF Then
        UserBelongsToRoyalArmy = False
        Exit Function

    End If

    UserBelongsToRoyalArmy = CBool(User_Database.Database_RecordSet!pertenece_real)
    Set User_Database.Database_RecordSet = Nothing
    
    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
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
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "SELECT pertenece_caos FROM personaje WHERE UPPER(name) = '" & UCase$(UserName) & "' AND deleted = FALSE;"

    Set User_Database.Database_RecordSet = User_Database.Database_Connection.Execute(query)

    If User_Database.Database_RecordSet.BOF Or User_Database.Database_RecordSet.EOF Then
        UserBelongsToChaosLegion = False
        Exit Function

    End If

    UserBelongsToChaosLegion = CBool(User_Database.Database_RecordSet!pertenece_caos)
    Set User_Database.Database_RecordSet = Nothing
    
    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
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
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "SELECT level FROM personaje WHERE UPPER(name) = '" & UCase$(UserName) & "';"

    Set User_Database.Database_RecordSet = User_Database.Database_Connection.Execute(query)

    If User_Database.Database_RecordSet.BOF Or User_Database.Database_RecordSet.EOF Then
        GetUserLevel = 0
        Exit Function

    End If

    GetUserLevel = CByte(User_Database.Database_RecordSet!level)
    Set User_Database.Database_RecordSet = Nothing
    
    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
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
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "SELECT rep_average FROM personaje WHERE UPPER(name) = '" & UCase$(UserName) & "';"

    Set User_Database.Database_RecordSet = User_Database.Database_Connection.Execute(query)

    If User_Database.Database_RecordSet.BOF Or User_Database.Database_RecordSet.EOF Then
        GetUserPromedio = 0
        Exit Function

    End If

    GetUserPromedio = CLng(User_Database.Database_RecordSet!rep_average)
    Set User_Database.Database_RecordSet = Nothing

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
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
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "SELECT reenlistadas FROM personaje WHERE UPPER(name) = '" & UCase$(UserName) & "';"

    Set User_Database.Database_RecordSet = User_Database.Database_Connection.Execute(query)

    If User_Database.Database_RecordSet.BOF Or User_Database.Database_RecordSet.EOF Then
        GetUserReenlists = 0
        Exit Function

    End If

    GetUserReenlists = CByte(User_Database.Database_RecordSet!Reenlistadas)
    Set User_Database.Database_RecordSet = Nothing
    
    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
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
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "UPDATE personaje SET "
    query = query & "reenlistadas = " & Reenlists & " "
    query = query & "WHERE UPPER(name) = '" & UCase$(UserName) & "';"

    User_Database.Database_Connection.Execute (query)

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
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
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If
    
        query = "SELECT level, exp, elu, min_sta, max_sta, min_hp, max_hp, min_man, max_man, min_hit, max_hit, gold FROM personaje WHERE UPPER(name) = '" & UCase$(UserName) & "';"

        Set User_Database.Database_RecordSet = User_Database.Database_Connection.Execute(query)

        If User_Database.Database_RecordSet.BOF Or User_Database.Database_RecordSet.EOF Then
            Call WriteConsoleMsg(sendIndex, "Pj Inexistente", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        Call WriteConsoleMsg(sendIndex, "Nivel: " & User_Database.Database_RecordSet!level & "  EXP: " & User_Database.Database_RecordSet!Exp & "/" & User_Database.Database_RecordSet!ELU, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Energia: " & User_Database.Database_RecordSet!min_sta & "/" & User_Database.Database_RecordSet!max_sta, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Salud: " & User_Database.Database_RecordSet!min_hp & "/" & User_Database.Database_RecordSet!max_hp, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Mana: " & User_Database.Database_RecordSet!min_man & "/" & User_Database.Database_RecordSet!max_man, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Golpe: " & User_Database.Database_RecordSet!min_hit & "/" & User_Database.Database_RecordSet!max_hit, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Oro: " & User_Database.Database_RecordSet!Gold, FontTypeNames.FONTTYPE_INFO)

        Set User_Database.Database_RecordSet = Nothing
        
        #If DBConexionUnica = 0 Then
            Call User_Database.Database_Close
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
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If
    
        query = "SELECT killed_npcs, killed_users, ciudadanos_matados, criminales_matados, class_id, genre_id, race_id FROM personaje WHERE UPPER(name) = '" & UCase$(UserName) & "';"

        Set User_Database.Database_RecordSet = User_Database.Database_Connection.Execute(query)

        If User_Database.Database_RecordSet.BOF Or User_Database.Database_RecordSet.EOF Then
            Call WriteConsoleMsg(sendIndex, "Pj Inexistente", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        Call WriteConsoleMsg(sendIndex, "Pj: " & UserName, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "CiudadanosMatados: " & User_Database.Database_RecordSet!ciudadanos_matados & ", CriminalesMatados: " & User_Database.Database_RecordSet!criminales_matados & ", UsuariosMatados: " & User_Database.Database_RecordSet!killed_users, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "NPCs muertos: " & User_Database.Database_RecordSet!killed_npcs, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Clase: " & ListaClases(User_Database.Database_RecordSet!class_id), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Genero: " & IIf(CByte(User_Database.Database_RecordSet!ciudadanos_matados) = eGenero.Hombre, "Hombre", "Mujer"), FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Raza: " & ListaRazas(User_Database.Database_RecordSet!race_id), FontTypeNames.FONTTYPE_INFO)

        Set User_Database.Database_RecordSet = Nothing
        
    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
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
            Call User_Database.Database_Connect
        #Else
            'Si perdimos la conexion reconectamos
            If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
        #End If

        query = "SELECT bank_gold FROM personaje WHERE UPPER(name) = '" & UCase$(UserName) & "';"

        Set User_Database.Database_RecordSet = User_Database.Database_Connection.Execute(query)

        If User_Database.Database_RecordSet.BOF Or User_Database.Database_RecordSet.EOF Then
            Call WriteConsoleMsg(sendIndex, "Pj Inexistente", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        Call WriteConsoleMsg(sendIndex, "Pj: " & UserName, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Oro en banco: " & User_Database.Database_RecordSet!bank_gold, FontTypeNames.FONTTYPE_INFO)

        Set User_Database.Database_RecordSet = Nothing
        
        #If DBConexionUnica = 0 Then
            Call User_Database.Database_Close
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

    Dim query   As String
    Dim LoopC   As Byte

    Dim ObjInd  As Long

    If Not PersonajeExiste(UserName) Then
        Call WriteConsoleMsg(sendIndex, "Pj Inexistente", FontTypeNames.FONTTYPE_INFO)
    Else
        #If DBConexionUnica = 0 Then
            Call User_Database.Database_Connect
        #Else
            'Si perdimos la conexion reconectamos
            If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
        #End If

        query = "SELECT "

        For LoopC = 1 To MAX_INVENTORY_SLOTS
            query = query & "item_id" & LoopC & ", amount" & LoopC
            If LoopC < MAX_INVENTORY_SLOTS Then query = query & ", "
        Next LoopC

        query = query & " FROM inventario_items WHERE user_id = (SELECT id FROM personaje WHERE UPPER(name) = '" & UCase$(UserName) & "')"

        Set User_Database.Database_RecordSet = User_Database.Database_Connection.Execute(query)

        If Not User_Database.Database_RecordSet.RecordCount = 0 Then
            User_Database.Database_RecordSet.MoveFirst

            While Not User_Database.Database_RecordSet.EOF

                ObjInd = val(User_Database.Database_RecordSet!item_id)

                If ObjInd > 0 Then
                    Call WriteConsoleMsg(sendIndex, "Objeto " & User_Database.Database_RecordSet!Number & " " & ObjData(ObjInd).Name & " Cantidad:" & User_Database.Database_RecordSet!Amount, FontTypeNames.FONTTYPE_INFO)

                End If

                User_Database.Database_RecordSet.MoveNext
            Wend
        Else
            Call WriteConsoleMsg(sendIndex, "Pj Inexistente", FontTypeNames.FONTTYPE_INFO)

        End If

        Set User_Database.Database_RecordSet = Nothing
        
        #If DBConexionUnica = 0 Then
            Call User_Database.Database_Close
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

    Dim query   As String
    Dim LoopC   As Byte

    Dim ObjInd As Long

    If Not PersonajeExiste(UserName) Then
        Call WriteConsoleMsg(sendIndex, "Pj Inexistente", FontTypeNames.FONTTYPE_INFO)
    Else
        #If DBConexionUnica = 0 Then
            Call User_Database.Database_Connect
        #Else
            'Si perdimos la conexion reconectamos
            If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
        #End If
        
        query = "SELECT "

        For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
            query = query & "item_id" & LoopC & ", amount" & LoopC
            If LoopC < MAX_BANCOINVENTORY_SLOTS Then query = query & ", "
        Next LoopC
        
        query = query & " FROM banco_items WHERE user_id = (SELECT id FROM personaje WHERE UPPER(name) = '" & UCase$(UserName) & "')"

        Set User_Database.Database_RecordSet = User_Database.Database_Connection.Execute(query)

        If Not User_Database.Database_RecordSet.RecordCount = 0 Then
            User_Database.Database_RecordSet.MoveFirst

            While Not User_Database.Database_RecordSet.EOF

                ObjInd = val(User_Database.Database_RecordSet!item_id)

                If ObjInd > 0 Then
                    Call WriteConsoleMsg(sendIndex, "Objeto " & User_Database.Database_RecordSet!Number & " " & ObjData(ObjInd).Name & " Cantidad:" & User_Database.Database_RecordSet!Amount, FontTypeNames.FONTTYPE_INFO)

                End If

                User_Database.Database_RecordSet.MoveNext
            Wend
        Else
            Call WriteConsoleMsg(sendIndex, "Pj Inexistente", FontTypeNames.FONTTYPE_INFO)

        End If

        Set User_Database.Database_RecordSet = Nothing
        
        #If DBConexionUnica = 0 Then
            Call User_Database.Database_Close
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
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "SELECT race_id, class_id, genre_id, level, gold, bank_gold, rep_average, guild_requests_history, guild_index, guild_member_history, pertenece_real, pertenece_caos, ciudadanos_matados, criminales_matados FROM personaje WHERE UPPER(name) = '" & UCase$(UserName) & "';"

    Set User_Database.Database_RecordSet = User_Database.Database_Connection.Execute(query)

    If User_Database.Database_RecordSet.BOF Or User_Database.Database_RecordSet.EOF Then
        Call WriteConsoleMsg(UserIndex, "Pj Inexistente", FontTypeNames.FONTTYPE_INFO)
        Exit Sub

    End If

    ' Get the character's current guild
    GuildActual = SanitizeNullValue(User_Database.Database_RecordSet!Guild_Index, 0)

    If GuildActual > 0 And GuildActual <= CANTIDADDECLANES Then
        gName = "<" & GuildName(GuildActual) & ">"
    Else
        gName = "Ninguno"

    End If

    'Get previous guilds
    Miembro = SanitizeNullValue(User_Database.Database_RecordSet!guild_member_history, vbNullString)

    If Len(Miembro) > 400 Then
        Miembro = ".." & Right$(Miembro, 400)

    End If

    Call Protocol.WriteCharacterInfo(UserIndex, UserName, User_Database.Database_RecordSet!race_id, User_Database.Database_RecordSet!class_id, User_Database.Database_RecordSet!genre_id, User_Database.Database_RecordSet!level, User_Database.Database_RecordSet!Gold, User_Database.Database_RecordSet!bank_gold, User_Database.Database_RecordSet!rep_average, SanitizeNullValue(User_Database.Database_RecordSet!guild_requests_history, vbNullString), gName, Miembro, User_Database.Database_RecordSet!pertenece_real, User_Database.Database_RecordSet!pertenece_caos, User_Database.Database_RecordSet!ciudadanos_matados, User_Database.Database_RecordSet!criminales_matados)

#If DBConexionUnica = 0 Then
    Call User_Database.Database_Close
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
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "SELECT guild_member_history FROM personaje WHERE UPPER(name) = '" & UCase$(UserName) & "';"

    Set User_Database.Database_RecordSet = User_Database.Database_Connection.Execute(query)

    If User_Database.Database_RecordSet.BOF Or User_Database.Database_RecordSet.EOF Then
        GetUserGuildMemberDatabase = vbNullString
        Exit Function

    End If

    GetUserGuildMemberDatabase = SanitizeNullValue(User_Database.Database_RecordSet!guild_member_history, vbNullString)
    Set User_Database.Database_RecordSet = Nothing
    
    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
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
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "SELECT guild_aspirant_index FROM personaje WHERE UPPER(name) = '" & UCase$(UserName) & "';"

    Set User_Database.Database_RecordSet = User_Database.Database_Connection.Execute(query)

    If User_Database.Database_RecordSet.BOF Or User_Database.Database_RecordSet.EOF Then
        GetUserGuildAspirantDatabase = 0
        Exit Function

    End If

    GetUserGuildAspirantDatabase = SanitizeNullValue(User_Database.Database_RecordSet!guild_aspirant_index, 0)
    Set User_Database.Database_RecordSet = Nothing
    
#If DBConexionUnica = 0 Then
    Call User_Database.Database_Close
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
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "SELECT guild_rejected_because FROM personaje WHERE UPPER(name) = '" & UCase$(UserName) & "';"

    Set User_Database.Database_RecordSet = User_Database.Database_Connection.Execute(query)

    If User_Database.Database_RecordSet.BOF Or User_Database.Database_RecordSet.EOF Then
        GetUserGuildRejectionReasonDatabase = vbNullString
        Exit Function

    End If

    GetUserGuildRejectionReasonDatabase = SanitizeNullValue(User_Database.Database_RecordSet!guild_rejected_because, vbNullString)
    Set User_Database.Database_RecordSet = Nothing
    
#If DBConexionUnica = 0 Then
    Call User_Database.Database_Close
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
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "SELECT guild_requests_history FROM personaje WHERE UPPER(name) = '" & UCase$(UserName) & "';"

    Set User_Database.Database_RecordSet = User_Database.Database_Connection.Execute(query)

    If User_Database.Database_RecordSet.BOF Or User_Database.Database_RecordSet.EOF Then
        GetUserGuildPedidosDatabase = vbNullString
        Exit Function

    End If

    GetUserGuildPedidosDatabase = SanitizeNullValue(User_Database.Database_RecordSet!guild_requests_history, vbNullString)
    Set User_Database.Database_RecordSet = Nothing
    
#If DBConexionUnica = 0 Then
    Call User_Database.Database_Close
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
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "UPDATE personaje SET "
    query = query & "guild_rejected_because = '" & Reason & "' "
    query = query & "WHERE UPPER(name) = '" & UCase$(UserName) & "';"

    User_Database.Database_Connection.Execute (query)
    
#If DBConexionUnica = 0 Then
    Call User_Database.Database_Close
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
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "UPDATE personaje SET "
    query = query & "guild_index = " & GuildIndex & " "
    query = query & "WHERE UPPER(name) = '" & UCase$(UserName) & "';"

    User_Database.Database_Connection.Execute (query)
    
#If DBConexionUnica = 0 Then
    Call User_Database.Database_Close
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
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "UPDATE personaje SET "
    query = query & "guild_aspirant_index = " & AspirantIndex & " "
    query = query & "WHERE UPPER(name) = '" & UCase$(UserName) & "';"

    User_Database.Database_Connection.Execute (query)

#If DBConexionUnica = 0 Then
    Call User_Database.Database_Close
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
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "UPDATE personaje SET "
    query = query & "guild_member_history = '" & guilds & "' "
    query = query & "WHERE UPPER(name) = '" & UCase$(UserName) & "';"

    User_Database.Database_Connection.Execute (query)

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
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
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "UPDATE personaje SET "
    query = query & "guild_requests_history = '" & Pedidos & "' "
    query = query & "WHERE UPPER(name) = '" & UCase$(UserName) & "';"

    User_Database.Database_Connection.Execute (query)

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
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
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    query = "UPDATE cuentas SET "
    query = query & "date_last_login = NOW(), "
    query = query & "last_ip = '" & UserIP & "' "
    query = query & "WHERE UPPER(username) = '" & UCase$(UserName) & "';"

    User_Database.Database_Connection.Execute (query)

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
    #End If

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in SaveAccountLastLoginDatabase: " & UserName & ". " & Err.Number & " - " & Err.description)

End Sub

Public Function SanitizeNullValue(ByVal Value As Variant, _
                                  ByVal defaultValue As Variant) As Variant
    SanitizeNullValue = IIf(IsNull(Value), defaultValue, Value)

End Function
