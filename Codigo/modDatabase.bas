Attribute VB_Name = "modDatabase"
Option Explicit

Private QueryBuilder As cStringBuilder

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
    '**************************************************
    'Author: Lorwik
    'Last Modification: 07/04/2021
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

    query = "INSERT INTO personaje SET name = (?), cuenta_id = (?), level = (?), exp = (?), elu = (?), genre_id = (?), race_id = (?), class_id = (?), "
    query = query & "home_id = (?), description = (?), gold = (?), elo = (?), "
    query = query & "pos_map = (?), pos_x = (?), pos_y = (?), body_id = (?), head_id = (?), weapon_id = (?), helmet_id = (?), shield_id = (?), "
    query = query & "items_amount = (?), slot_armour = (?), slot_weapon = (?), min_hp = (?), max_hp = (?), min_man = (?), max_man = (?), "
    query = query & "min_sta = (?), max_sta = (?), min_ham = (?), max_ham = (?), min_sed = (?), max_sed = (?), min_hit = (?), max_hit = (?), "
    query = query & "rep_noble = (?), rep_plebe = (?), rep_average = (?), profesionA = (?), ProfesionB = (?), levelPVP = (?), expPVP = (?), eluPVP = (?), origbody_id = (?), orighead_id = (?)"

    With UserList(UserIndex)

        Call User_Database.MakeQuery(query, True, .Name, .AccountInfo.ID, .Stats.ELV, .Stats.Exp, .Stats.ELU, .Genero, .Raza, .clase, .Hogar, .Desc, .Stats.Gld, _
                                    .Stats.ELO, .Pos.Map, .Pos.X, .Pos.Y, .Char.body, .Char.Head, .Char.WeaponAnim, .Char.CascoAnim, .Char.ShieldAnim, .Invent.NroItems, .Invent.ArmourEqpSlot, _
                                    .Invent.WeaponEqpSlot, .Stats.MinHp, .Stats.MaxHp, .Stats.MinMAN, .Stats.MaxMAN, .Stats.MinSta, .Stats.MaxSta, .Stats.MinHam, .Stats.MaxHam, _
                                    .Stats.MinAGU, .Stats.MaxAGU, .Stats.MinHIT, .Stats.MaxHIT, .Reputacion.NobleRep, .Reputacion.PlebeRep, .Reputacion.Promedio, _
                                    .Profesion(0).Profesion, .Profesion(1).Profesion, .Stats.ELVPVP, .Stats.ExpPVP, .Stats.ELUPVP, .OrigChar.body, .OrigChar.Head)
        
        
        'Obtenemos el ID del usuario
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
    '**************************************************
    'Author: Lorwik
    'Last Modification: 07/04/2021
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
        query = "UPDATE personaje SET name = (?), level = (?), exp = (?), elu = (?), genre_id = (?), race_id = (?), class_id = (?), home_id = (?), description = (?), "
        query = query & "gold = (?), bank_gold = (?), elo = (?), pet_amount = (?), pos_map = (?), pos_x = (?), pos_y = (?), last_map = (?), "
        query = query & "body_id = (?), head_id = (?), weapon_id = (?), helmet_id = (?), shield_id = (?), aura_id = (?), aura_color = (?), heading = (?), items_amount = (?), "
        query = query & "slot_armour = (?), slot_weapon = (?), slot_helmet = (?), slot_shield = (?), slot_ammo = (?), slot_ship = (?), slot_ring = (?), slot_bag = (?), "
        query = query & "min_hp = (?), max_hp = (?), min_man = (?), max_man = (?), min_sta = (?), max_sta = (?), min_ham = (?), max_ham = (?), min_sed = (?), max_sed = (?), min_hit = (?), max_hit = (?), "
        query = query & "killed_npcs = (?), killed_users = (?), rep_asesino = (?), rep_bandido = (?), rep_burgues = (?), rep_ladron = (?), rep_noble = (?), rep_plebe = (?), rep_average = (?), "
        query = query & "is_naked = (?), is_poisoned = (?), is_incinerado = (?), is_hidden = (?), is_hungry = (?), is_thirsty = (?), is_ban = (?), is_dead = (?), is_sailing = (?), is_paralyzed = (?), "
        query = query & "counter_pena = (?), pertenece_consejo_real = (?), pertenece_consejo_caos = (?), pertenece_real = (?), pertenece_caos = (?), ciudadanos_matados = (?), criminales_matados = (?), "
        query = query & "recibio_armadura_real = (?), recibio_armadura_caos = (?), recibio_exp_real = (?), recibio_exp_caos = (?), recompensas_real = (?), recompensas_caos = (?), "
        query = query & "reenlistadas = (?), fecha_ingreso = (?), nivel_ingreso = (?), matados_ingreso = (?), siguiente_recompensa = (?), guild_index = (?), is_global = (?), profesionA = (?), profesionB = (?), "
        query = query & "modocombate = (?), seguro = (?), levelPVP = (?), expPVP = (?), eluPVP = (?) WHERE id = (?)"

    With UserList(UserIndex)
            Call User_Database.MakeQuery(query, True, .Name, .Stats.ELV, .Stats.Exp, .Stats.ELU, .Genero, .Raza, .clase, .Hogar, .Desc, .Stats.Gld, .Stats.Banco, .Stats.ELO, .NroMascotas, _
                                        .Pos.Map, .Pos.X, .Pos.Y, .flags.lastMap, .Char.body, .Char.Head, .Char.WeaponAnim, .Char.CascoAnim, .Char.ShieldAnim, .Char.AuraAnim, .Char.AuraColor, .Char.Heading, .Invent.NroItems, _
                                        .Invent.ArmourEqpSlot, .Invent.WeaponEqpSlot, .Invent.CascoEqpSlot, .Invent.EscudoEqpSlot, .Invent.MunicionEqpSlot, .Invent.BarcoSlot, .Invent.AnilloEqpSlot, .Invent.MochilaEqpSlot, _
                                        .Stats.MinHp, .Stats.MaxHp, .Stats.MinMAN, .Stats.MaxMAN, .Stats.MinSta, .Stats.MaxSta, .Stats.MinHam, .Stats.MaxHam, .Stats.MinAGU, .Stats.MaxAGU, .Stats.MinHIT, .Stats.MaxHIT, _
                                        .Stats.NPCsMuertos, .Stats.UsuariosMatados, .Reputacion.AsesinoRep, .Reputacion.BandidoRep, .Reputacion.BurguesRep, .Reputacion.LadronesRep, .Reputacion.NobleRep, .Reputacion.PlebeRep, _
                                        .Reputacion.Promedio, .flags.Desnudo, .flags.Envenenado, .flags.Incinerado, .flags.Escondido, .flags.Hambre, .flags.Sed, .flags.Ban, .flags.Muerto, .flags.Navegando, .flags.Paralizado, _
                                        .Counters.Pena, (.flags.Privilegios And PlayerType.RoyalCouncil), (.flags.Privilegios And PlayerType.ChaosCouncil), .Faccion.ArmadaReal, .Faccion.FuerzasCaos, .Faccion.CiudadanosMatados, _
                                        .Faccion.CriminalesMatados, .Faccion.RecibioArmaduraReal, .Faccion.RecibioArmaduraCaos, .Faccion.RecibioExpInicialReal, .Faccion.RecibioExpInicialCaos, .Faccion.RecompensasReal, _
                                        .Faccion.RecompensasCaos, .Faccion.Reenlistadas, .Faccion.FechaIngreso, .Faccion.NivelIngreso, .Faccion.MatadosIngreso, .Faccion.NextRecompensa, .GuildIndex, .flags.Global, .Profesion(0).Profesion, _
                                        .Profesion(1).Profesion, IIf(.flags.ModoCombate = True, "1", "0"), IIf(.flags.Seguro = True, "1", "0"), .Stats.ELVPVP, .Stats.ExpPVP, .Stats.ELUPVP, .ID)
                                        

        '*******************************************************************
        'Hechizos
        '*******************************************************************
        
        query = "INSERT INTO spell (user_id, slot, spell_id) VALUES "
        
        For LoopC = 1 To MAXUSERHECHIZOS
            
            query = query & "("
            query = query & .ID & ", "
            query = query & LoopC & ", "
            query = query & .Stats.UserHechizos(LoopC) & ") "
            
            If LoopC < MAXUSERHECHIZOS Then query = query & ", "
            
        Next LoopC
        
        query = query & " ON DUPLICATE KEY UPDATE spell_id=VALUES(spell_id); "

        Call User_Database.Database_Connection.Execute(query)

        '*******************************************************************
        'Inventario
        '*******************************************************************
        
        query = "INSERT INTO inventario_items (user_id, slot, item_id, Amount, is_equipped) VALUES "
        
        For LoopC = 1 To MAX_INVENTORY_SLOTS
        
            query = query & "("
            query = query & .ID & ", "
            query = query & LoopC & ", "
            query = query & .Invent.Object(LoopC).ObjIndex & ", "
            query = query & .Invent.Object(LoopC).Amount & ", "
            query = query & .Invent.Object(LoopC).Equipped & ")"
            
            If LoopC < MAX_INVENTORY_SLOTS Then query = query & ", "
            
        Next LoopC
        
        query = query & " ON DUPLICATE KEY UPDATE item_id=VALUES(item_id), amount=VALUES(Amount), is_equipped=VALUES(is_equipped); "
        Call User_Database.Database_Connection.Execute(query)

        '*******************************************************************
        'Boveda
        '*******************************************************************
        
        query = "INSERT INTO banco_items (user_id, slot, item_id, Amount) VALUES "
        
        For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
            query = query & "("
            query = query & .ID & ", "
            query = query & LoopC & ", "
            query = query & .BancoInvent.Object(LoopC).ObjIndex & ", "
            query = query & .BancoInvent.Object(LoopC).Amount & ")"
            
            If LoopC < MAX_BANCOINVENTORY_SLOTS Then query = query & ", "
            
        Next LoopC
        
        query = query & " ON DUPLICATE KEY UPDATE item_id=VALUES(item_id), amount=VALUES(Amount); "
        
        Call User_Database.Database_Connection.Execute(query)

        '*******************************************************************
        'Skills
        '*******************************************************************
        query = "INSERT INTO skillpoint (user_id, skill_id, sk, exp, elu) VALUES "
        
        For LoopC = 1 To NUMSKILLS
            query = query & "("
            query = query & .ID & ", "
            query = query & LoopC & ", "
            query = query & .Stats.UserSkills(LoopC) & ", "
            query = query & .Stats.ExpSkills(LoopC) & ", "
            query = query & .Stats.EluSkills(LoopC) & ")"
            
            If LoopC < NUMSKILLS Then query = query & ", "

        Next LoopC
        
        query = query & " ON DUPLICATE KEY UPDATE sk=VALUES(sk), exp=VALUES(exp), elu=VALUES(elu); "
        
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
        
        query = "INSERT INTO pet (user_id, slot, pet_id) VALUES "
        
        For LoopC = 1 To MAXMASCOTAS
            
            query = query & "("
            query = query & .ID & ", "
            query = query & LoopC & ", "
            
            If .MascotasIndex(LoopC) > 0 Then
            
                If Npclist(.MascotasIndex(LoopC)).Contadores.TiempoExistencia = 0 Then
                    petType = .MascotasType(LoopC)
                    
                Else
                    petType = 0

                End If

            Else
                petType = .MascotasType(LoopC)

            End If

            query = query & petType & ")"
            
            If LoopC < MAXMASCOTAS Then query = query & ", "
        Next LoopC
        
        query = query & " ON DUPLICATE KEY UPDATE pet_id=VALUES(pet_id); "

        Call User_Database.Database_Connection.Execute(query)
        
        '*******************************************************************
        'Amigos
        '*******************************************************************
        query = "INSERT INTO amigos (user_id, slot, amigo, ignorado) VALUES "
        
        For LoopC = 1 To MAXAMIGOS
            query = query & "("
            query = query & .ID & ", "
            query = query & LoopC & ", "
            query = query & "'" & .Amigos(LoopC).Nombre & "', "
            query = query & .Amigos(LoopC).Ignorado & ")"
            
            If LoopC < MAXAMIGOS Then query = query & ", "
        Next LoopC
        
        query = query & " ON DUPLICATE KEY UPDATE amigo=VALUES(amigo), ignorado=VALUES(ignorado); "
            
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

    Dim LoopC As Integer

    Dim j     As Integer

    Dim tmpst As String
    
    Set QueryBuilder = New cStringBuilder
    
    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Connect
    #Else

        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If
    
    'Basic user data
    With UserList(UserIndex)
    
        '¿El usuario tiene alguna quest que guardar?
        If .QuestStats.TotalQuest = 0 Then Exit Sub

        QueryBuilder.Append "INSERT INTO quest (user_id, quest_id, npcs, target, fechafin, estado) VALUES "
        
        For LoopC = 1 To .QuestStats.TotalQuest
            
            QueryBuilder.Append "("
            QueryBuilder.Append .ID & ", "
            QueryBuilder.Append .QuestStats.Quests(LoopC).QuestIndex & ", '"
                
            'NPC
            If .QuestStats.Quests(LoopC).QuestIndex > 0 Then
                tmpst = QuestList(.QuestStats.Quests(LoopC).QuestIndex).RequiredNPCs
    
                If tmpst Then
    
                    For j = 1 To tmpst
                        QueryBuilder.Append CStr(.QuestStats.Quests(LoopC).NPCsKilled(j))

                        If j < tmpst Then QueryBuilder.Append "-"
                    Next j
    
                End If
    
            End If
                
            QueryBuilder.Append "', '"
          
            'TARGETs
            If .QuestStats.Quests(LoopC).QuestIndex > 0 Then
              
                tmpst = QuestList(.QuestStats.Quests(LoopC).QuestIndex).RequiredTargetNPCs
                  
                For j = 1 To tmpst
                    QueryBuilder.Append CStr(.QuestStats.Quests(LoopC).NPCsTarget(j))

                    If j < tmpst Then QueryBuilder.Append "-"
                Next j
          
            End If
          
            QueryBuilder.Append "', '"
            QueryBuilder.Append Format(.QuestStats.Quests(LoopC).fechaFin, "yyyy/MM/dd") & "', "
            QueryBuilder.Append CInt(.QuestStats.Quests(LoopC).QuestStatus)
            QueryBuilder.Append ")"
                
            If LoopC < .QuestStats.TotalQuest Then QueryBuilder.Append ", "
    
        Next LoopC
        
        QueryBuilder.Append " ON DUPLICATE KEY UPDATE estado = VALUES(estado), npcs = VALUES(npcs), target = VALUES(target); "

        Call User_Database.Database_Connection.Execute(QueryBuilder.toString)
        
        Set QueryBuilder = Nothing

    End With

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
    #End If

    Exit Sub

ErrorHandler:
    Call LogDatabaseError("Unable to UPDATE personaje to Mysql Database: " & UserList(UserIndex).Name & ". " & Err.Number & " - " & Err.description)

End Sub

Sub LoadUserFromDatabase(ByVal UserIndex As Integer)
    '**************************************************
    'Author: Lorwik
    'Last Modification: 07/04/2021
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
        query = "SELECT *, DATE_FORMAT(fecha_ingreso, '%Y-%m-%d') as 'fecha_ingreso_format' FROM personaje WHERE UPPER(name) = (?)"
        
        If Not User_Database.MakeQuery(query, False, UCase$(.Name)) Then Exit Sub

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
        .Stats.ELO = User_Database.Database_RecordSet!ELO
        .NroMascotas = User_Database.Database_RecordSet!pet_amount
        .Pos.Map = User_Database.Database_RecordSet!pos_map
        .Pos.X = User_Database.Database_RecordSet!pos_x
        .Pos.Y = User_Database.Database_RecordSet!pos_y
        .flags.lastMap = User_Database.Database_RecordSet!last_map
        .OrigChar.body = User_Database.Database_RecordSet!origbody_id
        .OrigChar.Head = User_Database.Database_RecordSet!orighead_id
        .OrigChar.WeaponAnim = User_Database.Database_RecordSet!weapon_id
        .OrigChar.CascoAnim = User_Database.Database_RecordSet!helmet_id
        .OrigChar.ShieldAnim = User_Database.Database_RecordSet!shield_id
        .OrigChar.Heading = User_Database.Database_RecordSet!Heading
        .OrigChar.AuraAnim = User_Database.Database_RecordSet!Aura_id
        .OrigChar.AuraColor = User_Database.Database_RecordSet!Aura_color
        .Char.body = User_Database.Database_RecordSet!body_id
        .Char.Head = User_Database.Database_RecordSet!head_id
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
        .Profesion(0).Profesion = User_Database.Database_RecordSet!profesionA
        .Profesion(1).Profesion = User_Database.Database_RecordSet!profesionB
        .flags.ModoCombate = User_Database.Database_RecordSet!ModoCombate
        .flags.Seguro = User_Database.Database_RecordSet!Seguro
        .Stats.ELVPVP = User_Database.Database_RecordSet!levelPVP
        .Stats.ExpPVP = User_Database.Database_RecordSet!ExpPVP
        .Stats.ELUPVP = User_Database.Database_RecordSet!ELUPVP
        
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
        Call User_Database.MakeQuery("SELECT * FROM atributos WHERE user_id = (?)", False, .ID)
    
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
        If User_Database.MakeQuery("SELECT * FROM spell WHERE user_id = (?)", False, .ID) Then

            User_Database.Database_RecordSet.MoveFirst

            While Not User_Database.Database_RecordSet.EOF
            
                LoopC = User_Database.Database_RecordSet!Slot
                
                .Stats.UserHechizos(LoopC) = User_Database.Database_RecordSet!spell_id
                
                User_Database.Database_RecordSet.MoveNext
                
            Wend

        End If

        Set User_Database.Database_RecordSet = Nothing

        '*******************************************************************
        'Mascotas
        '*******************************************************************
        If User_Database.MakeQuery("SELECT * FROM pet WHERE user_id = (?)", False, .ID) Then
            User_Database.Database_RecordSet.MoveFirst

            While Not User_Database.Database_RecordSet.EOF
            
                LoopC = User_Database.Database_RecordSet!Slot
                
                .MascotasType(LoopC) = User_Database.Database_RecordSet!pet_id
                
                User_Database.Database_RecordSet.MoveNext
                
            Wend

        End If

        Set User_Database.Database_RecordSet = Nothing

        '*******************************************************************
        'Inventario
        '*******************************************************************

        If User_Database.MakeQuery("SELECT * FROM inventario_items WHERE user_id = (?)", False, .ID) Then
            User_Database.Database_RecordSet.MoveFirst
                
            While Not User_Database.Database_RecordSet.EOF
            
                LoopC = User_Database.Database_RecordSet!Slot
                
                .Invent.Object(LoopC).ObjIndex = User_Database.Database_RecordSet!item_id
                .Invent.Object(LoopC).Amount = User_Database.Database_RecordSet!Amount
                .Invent.Object(LoopC).Equipped = User_Database.Database_RecordSet!is_equipped
                
                 User_Database.Database_RecordSet.MoveNext
            Wend
                
        End If

        Set User_Database.Database_RecordSet = Nothing

        '*******************************************************************
        'Boveda
        '*******************************************************************
        If User_Database.MakeQuery("SELECT * FROM banco_items WHERE user_id = (?)", False, .ID) Then
            User_Database.Database_RecordSet.MoveFirst
                
            While Not User_Database.Database_RecordSet.EOF
            
                LoopC = User_Database.Database_RecordSet!Slot
            
                .BancoInvent.Object(LoopC).ObjIndex = User_Database.Database_RecordSet!item_id
                .BancoInvent.Object(LoopC).Amount = User_Database.Database_RecordSet!Amount
                
                User_Database.Database_RecordSet.MoveNext
            Wend
                
        End If

        Set User_Database.Database_RecordSet = Nothing

        '*******************************************************************
        'Skills
        '*******************************************************************
        If User_Database.MakeQuery("SELECT * FROM skillpoint WHERE user_id = (?)", False, .ID) Then
            User_Database.Database_RecordSet.MoveFirst

            While Not User_Database.Database_RecordSet.EOF
            
                LoopC = User_Database.Database_RecordSet!skill_id
                
                .Stats.UserSkills(LoopC) = User_Database.Database_RecordSet!sk
                .Stats.ExpSkills(LoopC) = User_Database.Database_RecordSet!Exp
                .Stats.EluSkills(LoopC) = User_Database.Database_RecordSet!ELU
                
                User_Database.Database_RecordSet.MoveNext
            Wend

        End If

        Set User_Database.Database_RecordSet = Nothing
        
        '*******************************************************************
        'Profesion primaria
        '*******************************************************************
        Call User_Database.MakeQuery("SELECT * FROM profesion_primaria WHERE user_id = (?)", False, .ID)

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
        Call User_Database.MakeQuery("SELECT * FROM profesion_secundaria WHERE user_id = (?)", False, .ID)

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
        If User_Database.MakeQuery("SELECT * FROM amigos WHERE user_id = (?)", False, .ID) Then
            User_Database.Database_RecordSet.MoveFirst

            While Not User_Database.Database_RecordSet.EOF
            
                LoopC = User_Database.Database_RecordSet!Slot
                
                .Amigos(LoopC).Nombre = User_Database.Database_RecordSet!Amigo
                .Amigos(LoopC).Ignorado = User_Database.Database_RecordSet!Ignorado
                
                User_Database.Database_RecordSet.MoveNext
            Wend

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

    Dim query         As String

    Dim Fields()      As String

    Dim j             As Integer

    Dim tmpint        As Integer

    Dim questID       As Integer

    Dim NPCRequeridos As String
    
    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Connect
    #Else

        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    With UserList(UserIndex).QuestStats

        If User_Database.MakeQuery("SELECT * FROM quest WHERE user_id = (?)", False, UserList(UserIndex).ID) Then
        
            User_Database.Database_RecordSet.MoveFirst
            
            While Not User_Database.Database_RecordSet.EOF
            
                .TotalQuest = .TotalQuest + 1
                
                ReDim Preserve .Quests(1 To .TotalQuest) As tUserQuest
            
                .Quests(.TotalQuest).QuestIndex = User_Database.Database_RecordSet!quest_id
                questID = .Quests(.TotalQuest).QuestIndex

                '¿La quest requiere matar NPC's?
                If QuestList(questID).RequiredNPCs Then
                    ReDim .Quests(.TotalQuest).NPCsKilled(1 To QuestList(questID).RequiredNPCs)
            
                    NPCRequeridos = User_Database.Database_RecordSet("npcs")
                    Fields = Split(NPCRequeridos, "-")
                        
                    For j = 1 To QuestList(questID).RequiredNPCs
                        .Quests(.TotalQuest).NPCsKilled(j) = val(Fields(j - 1))
                    Next j
 
                End If
                
                '¿La quest requiere hablar con NPC's?
                If QuestList(questID).RequiredTargetNPCs Then
                    ReDim .Quests(.TotalQuest).NPCsTarget(1 To QuestList(questID).RequiredTargetNPCs)
            
                    NPCRequeridos = User_Database.Database_RecordSet("target")
                    Fields = Split(NPCRequeridos, "-")
                        
                    For j = 1 To QuestList(questID).RequiredTargetNPCs
                        .Quests(.TotalQuest).NPCsTarget(j) = val(Fields(j - 1))
                    Next j
 
                End If
                
                .Quests(.TotalQuest).fechaFin = User_Database.Database_RecordSet("fechafin")
                .Quests(.TotalQuest).QuestStatus = CByte(User_Database.Database_RecordSet("estado"))
                
                'Clasificacion de quest por estado
                'Vamos guardando el slot donde se encuentra la quest en su array correspondiente
                Select Case .Quests(.TotalQuest).QuestStatus
                
                    Case eStatusQuest.NoAceptada
                        .nQuestLeave = .nQuestLeave + 1
                        
                        ReDim Preserve .QuestLeave(1 To .nQuestLeave) As Integer
                        .QuestLeave(.nQuestLeave) = .TotalQuest
                    
                    Case eStatusQuest.EnCurso
                        .nQuestCurso = .nQuestCurso + 1
                        
                        ReDim Preserve .QuestEnCurso(1 To .nQuestCurso) As Integer
                        .QuestEnCurso(.nQuestCurso) = .TotalQuest
                    
                    Case eStatusQuest.Terminada
                        .nQuestDone = .nQuestDone + 1
                        
                        ReDim Preserve .QuestDone(1 To .nQuestCurso) As Integer
                        .QuestDone(.nQuestDone) = .TotalQuest
                
                End Select
           
                User_Database.Database_RecordSet.MoveNext
                
            Wend
            
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

Public Function PersonajeExisteDatabase(ByVal username As String) As Boolean
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

    query = "SELECT id FROM personaje WHERE UPPER(name) = (?) AND deleted = FALSE;"

    If Not User_Database.MakeQuery(query, False, UCase$(username)) Then
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
    Call LogDatabaseError("Error in PersonajeExisteDatabase: " & username & ". " & Err.Number & " - " & Err.description)

End Function

Public Function BANCheckDatabase(ByVal username As String) As Boolean

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

    query = "SELECT is_ban FROM personaje WHERE UPPER(name) = (?)"

    If Not User_Database.MakeQuery(query, False, UCase$(username)) Then
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
    Call LogDatabaseError("Error in BANCheckDatabase: " & username & ". " & Err.Number & " - " & Err.description)

End Function

Public Function BanTimeCheck(ByVal username As String) As Date

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

    query = "SELECT ban_time FROM personaje WHERE UPPER(name) = (?)"

    If Not User_Database.MakeQuery(query, False, UCase$(username)) Then
        BanTimeCheck = False
        Exit Function

    End If

    BanTimeCheck = User_Database.Database_RecordSet!ban_time

    Set User_Database.Database_RecordSet = Nothing
    
    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
    #End If

    Exit Function

ErrorHandler:
    Call LogDatabaseError("Error in BanTimeCheck: " & username & ". " & Err.Number & " - " & Err.description)

End Function

Public Sub UnBanDatabase(ByVal username As String)

    '***************************************************
    'Author: Lorwik
    'Last Modification: 07/04/2021
    '***************************************************
    On Error GoTo ErrorHandler
    
    #If DBConexionUnica = 0 Then
        Call Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    Call User_Database.MakeQuery("UPDATE personaje SET is_ban = FALSE WHERE UPPER(name) = (?)", True, UCase$(username))

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
    #End If

    Exit Sub

ErrorHandler:
    Call LogDatabaseError("Error in UnBanDatabase: " & username & ". " & Err.Number & " - " & Err.description)

End Sub

Public Function GetUserGuildIndexDatabase(ByVal username As String) As Integer

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

    query = "SELECT guild_index FROM personaje WHERE UPPER(name) = (?)"
    
    If Not User_Database.MakeQuery(query, False, UCase$(username)) Then
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
    Call LogDatabaseError("Error in GetUserGuildIndexDatabase: " & username & ". " & Err.Number & " - " & Err.description)

End Function

Public Sub CopyUserDatabase(ByVal username As String, ByVal newName As String)

    '***************************************************
    'Author: Lorwik
    'Last Modification: 07/04/2021
    '***************************************************
    On Error GoTo ErrorHandler

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    Call User_Database.MakeQuery("UPDATE personaje SET name = (?) WHERE UPPER(name) = (?)", True, newName, UCase$(username))

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
    #End If

    Exit Sub

ErrorHandler:
    Call LogDatabaseError("Error in CopyUserDatabase: " & username & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub MarcarPjComoQueYaVotoDatabase(ByVal UserIndex As Integer, _
                                         ByVal NumeroEncuesta As Integer)

    '***************************************************
    'Author: Lorwik
    'Last Modification: 07/04/2021
    '***************************************************
    On Error GoTo ErrorHandler

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    Call User_Database.MakeQuery("UPDATE personaje SET votes_amount = (?) WHERE id = (?)", True, NumeroEncuesta, UserList(UserIndex).ID)
    
    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
    #End If

    Exit Sub

ErrorHandler:
    Call LogDatabaseError("Error in MarcarPjComoQueYaVotoDatabase: " & UserList(UserIndex).Name & ". " & Err.Number & " - " & Err.description)

End Sub

Public Function PersonajeCantidadVotosDatabase(ByVal username As String) As Integer

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

    query = "SELECT votes_amount FROM personaje WHERE UPPER(name) = (?)"
    
    If Not User_Database.MakeQuery(query, False, UCase$(username)) Then
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
    Call LogDatabaseError("Error in PersonajeCantidadVotosDatabase: " & username & ". " & Err.Number & " - " & Err.description)

End Function

Public Sub SaveBan(ByVal username As String, _
                           ByVal Reason As String, _
                           ByVal BannedBy As String, _
                           Optional ByVal Tiempo As Date = 0)

    '***************************************************
    'Author: Lorwik
    'Last Modification: 07/04/2021
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query     As String

    Dim cantPenas As Byte

    cantPenas = GetUserAmountOfPunishments(username)

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If

    Call User_Database.MakeQuery("UPDATE personaje SET is_ban = TRUE, ban_time = (?) WHERE UPPER(name) = (?)", True, Tiempo, UCase$(username))

    query = "INSERT INTO punishment SET user_id = (SELECT id FROM personaje WHERE UPPER(name) = (?)), number = (?), reason = (?)"
    Call User_Database.MakeQuery(query, True, UCase$(username), (cantPenas + 1), BannedBy & ": BAN POR " & LCase$(Reason) & " " & Date & " " & time)

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
    #End If

    Exit Sub

ErrorHandler:
    Call LogDatabaseError("Error in SaveBan: " & username & ". " & Err.Number & " - " & Err.description)

End Sub

Public Function GetUserAmountOfPunishments(ByVal username As String) As Integer

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

    query = "SELECT COUNT(1) as punishments FROM punishment WHERE user_id = (SELECT id FROM personaje WHERE UPPER(name) = (?))"
    
    If Not User_Database.MakeQuery(query, False, UCase$(username)) Then
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
    Call LogDatabaseError("Error in GetUserAmountOfPunishments: " & username & ". " & Err.Number & " - " & Err.description)

End Function

Public Sub SendUserPunishments(ByVal UserIndex As Integer, _
                                       ByVal username As String, _
                                       ByVal Count As Integer)

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

    query = "SELECT * FROM punishment WHERE user_id = (SELECT id FROM personaje WHERE UPPER(name) = (?))"
    
    If Not User_Database.MakeQuery(query, False, UCase$(username)) Then
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
    Call LogDatabaseError("Error in SendUserPunishments: " & username & ". " & Err.Number & " - " & Err.description)

End Sub

Public Function GetUserPos(ByVal username As String) As String

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

    query = "SELECT pos_map, pos_x, pos_y FROM personaje WHERE UPPER(name) = (?)"
    
    If Not User_Database.MakeQuery(query, False, UCase$(username)) Then
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
    Call LogDatabaseError("Error in GetUserPos: " & username & ". " & Err.Number & " - " & Err.description)

End Function

Public Sub SaveUserPunishment(ByVal username As String, _
                                      ByVal Number As Integer, _
                                      ByVal Reason As String)

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

    query = "INSERT INTO punishment SET user_id = (SELECT id FROM personaje WHERE UPPER(name) = (?)), number = (?), reason = (?)"
    Call User_Database.MakeQuery(query, True, UCase$(username), Number, Reason)

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
    #End If

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in SaveUserPunishment: " & username & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub AlterUserPunishment(ByVal username As String, _
                                       ByVal Number As Integer, _
                                       ByVal Reason As String)

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

    query = "UPDATE punishment SET reason = (?) WHERE number = (?) AND user_id = (SELECT id FROM personaje WHERE UPPER(name) = (?)"
    Call User_Database.MakeQuery(query, True, Reason, Number, UCase$(username))

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
    #End If

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in AlterUserPunishment: " & username & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub ResetUserFacciones(ByVal username As String)

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

    query = "UPDATE personaje SET pertenece_real = FALSE, pertenece_caos = FALSE, ciudadanos_matados = 0, criminales_matados = FALSE, "
    query = query & "recibio_armadura_real = FALSE, recibio_armadura_caos = FALSE, recibio_exp_real = FALSE, recibio_exp_caos = FALSE, "
    query = query & "recompensas_real = 0, recompensas_caos = 0, reenlistadas = 0, fecha_ingreso = NULL, nivel_ingreso = NULL, "
    query = query & "matados_ingreso = NULL, siguiente_recompensa = NULL WHERE UPPER(name) = (?)"

    Call User_Database.MakeQuery(query, True, UCase$(username))

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
    #End If

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in ResetUserFacciones: " & username & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub KickUserCouncils(ByVal username As String)

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

    query = "UPDATE personaje SET pertenece_consejo_real = FALSE, pertenece_consejo_caos = FALSE WHERE UPPER(name) = (?)"
    Call User_Database.MakeQuery(query, True, UCase$(username))

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
    #End If

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in KickUserCouncils: " & username & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub KickUserFacciones(ByVal username As String)

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

    query = "UPDATE personaje SET pertenece_real = FALSE, pertenece_caos = FALSE WHERE UPPER(name) = (?)"
    Call User_Database.MakeQuery(query, True, UCase$(username))
    
    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
    #End If

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in KickUserFacciones: " & username & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub KickUserChaosLegion(ByVal username As String)

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

    query = "UPDATE personaje SET pertenece_caos = FALSE, reenlistadas = 200 WHERE UPPER(name) = (?)"
    Call User_Database.MakeQuery(query, True, UCase$(username))

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
    #End If

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in KickUserChaosLegion: " & username & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub KickUserRoyalArmy(ByVal username As String)

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

    query = "UPDATE personaje SET pertenece_real = FALSE, reenlistadas = 200 WHERE UPPER(name) = (?)"
    Call User_Database.MakeQuery(query, True, UCase$(username))

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
    #End If

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in KickUserRoyalArmy: " & username & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub UpdateUserLogged(ByVal username As String, ByVal Logged As Byte)

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

    query = "UPDATE personaje SET is_logged = " & IIf(Logged = 1, "TRUE", "FALSE") & " WHERE UPPER(name) = (?)"
    Call User_Database.MakeQuery(query, True, UCase$(username))

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
    #End If

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in UpdateUserLogged: " & username & ". " & Err.Number & " - " & Err.description)

End Sub

Public Function GetUserLastIps(ByVal username As String) As String

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

    query = "SELECT last_ip FROM cuentas WHERE id = (SELECT cuenta_id FROM personaje WHERE UPPER(name) = (?)"
    
    If Not User_Database.MakeQuery(query, False, UCase$(username)) Then
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
    Call LogDatabaseError("Error in GetUserLastIps: " & username & ". " & Err.Number & " - " & Err.description)

End Function

Public Function GetUserSkills(ByVal username As String) As String

    '***************************************************
    'Author: Lorwik
    'Last Modification: 07/04/2021
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

    query = "SELECT number, value FROM skillpoint WHERE user_id = (SELECT id FROM personaje WHERE UPPER(name) = (?))"
    
   If Not User_Database.MakeQuery(query, False, UCase$(username)) Then
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
    Call LogDatabaseError("Error in GetUserSkills: " & username & ". " & Err.Number & " - " & Err.description)

End Function

Public Function GetUserFreeSkills(ByVal username As String) As Integer

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

    query = "SELECT free_skillpoints FROM personaje WHERE UPPER(name) = (?)"
    If Not User_Database.MakeQuery(query, False, UCase$(username)) Then
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
    Call LogDatabaseError("Error in GetUserFreeSkills: " & username & ". " & Err.Number & " - " & Err.description)

End Function

Public Sub SaveUserTrainingTime(ByVal username As String, _
                                        ByVal trainingTime As Long)

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

    query = "UPDATE personaje SET counter_training = (?) WHERE UPPER(name) = (?)"
    Call User_Database.MakeQuery(query, True, trainingTime, UCase$(username))

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
    #End If

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in SaveUserTrainingTime: " & username & ". " & Err.Number & " - " & Err.description)

End Sub

Public Function GetUserTrainingTime(ByVal username As String) As Long

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

    query = "SELECT counter_training FROM personaje WHERE UPPER(name) = (?)"
    
    If Not User_Database.MakeQuery(query, False, UCase$(username)) Then
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
    Call LogDatabaseError("Error in GetUserTrainingTime: " & username & ". " & Err.Number & " - " & Err.description)

End Function

Public Function UserBelongsToRoyalArmy(ByVal username As String) As Boolean

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

    query = "SELECT pertenece_real FROM personaje WHERE UPPER(name) = (?) AND deleted = FALSE;"
    
    If Not User_Database.MakeQuery(query, False, UCase$(username)) Then
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
    Call LogDatabaseError("Error in UserBelongsToRoyalArmy: " & username & ". " & Err.Number & " - " & Err.description)

End Function

Public Function UserBelongsToChaosLegion(ByVal username As String) As Boolean

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

    query = "SELECT pertenece_caos FROM personaje WHERE UPPER(name) = (?) AND deleted = FALSE;"
    
    If Not User_Database.MakeQuery(query, False, UCase$(username)) Then
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
    Call LogDatabaseError("Error in UserBelongsToChaosLegion: " & username & ". " & Err.Number & " - " & Err.description)

End Function

Public Function GetUserLevel(ByVal username As String) As Byte

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

    query = "SELECT level FROM personaje WHERE UPPER(name) = (?)"
    
    If Not User_Database.MakeQuery(query, False, UCase$(username)) Then
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
    Call LogDatabaseError("Error in GetUserLevel: " & username & ". " & Err.Number & " - " & Err.description)

End Function

Public Function GetUserPromedio(ByVal username As String) As Long

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

    query = "SELECT rep_average FROM personaje WHERE UPPER(name) = (?)"
    
    If Not User_Database.MakeQuery(query, False, UCase$(username)) Then
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
    Call LogDatabaseError("Error in GetUserPromedio: " & username & ". " & Err.Number & " - " & Err.description)

End Function

Public Function GetUserReenlists(ByVal username As String) As Byte

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

    query = "SELECT reenlistadas FROM personaje WHERE UPPER(name) = (?)"
    
    If Not User_Database.MakeQuery(query, False, UCase$(username)) Then
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
    Call LogDatabaseError("Error in GetUserReenlists: " & username & ". " & Err.Number & " - " & Err.description)

End Function

Public Function GetAccountID(ByVal username As String) As Long

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

    If Not User_Database.MakeQuery("SELECT cuenta_id FROM personaje WHERE UPPER(name) = (?)", False, UCase$(username)) Then
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
    Call LogDatabaseError("Error in GetAccountID: " & username & ". " & Err.Number & " - " & Err.description)

End Function

Public Sub SaveUserReenlists(ByVal username As String, ByVal Reenlists As Byte)

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

    query = "UPDATE personaje SET reenlistadas = (?) WHERE UPPER(name) = (?)"

    Call User_Database.MakeQuery(query, True, Reenlists, UCase$(username))

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
    #End If

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in SaveUserReenlists: " & username & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub SendUserStatsTxtDatabase(ByVal sendIndex As Integer, ByVal username As String)

    '***************************************************
    'Author: Lorwik
    'Last Modification: 07/04/2021
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    If Not PersonajeExiste(username) Then
        Call WriteConsoleMsg(sendIndex, "Pj Inexistente", FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(sendIndex, "Estadisticas de: " & username, FontTypeNames.FONTTYPE_INFO)

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If
    
        query = "SELECT level, exp, elu, min_sta, max_sta, min_hp, max_hp, min_man, max_man, min_hit, max_hit, gold FROM personaje WHERE UPPER(name) = (?)"
        
        If Not User_Database.MakeQuery(query, False, UCase$(username)) Then
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
    Call LogDatabaseError("Error in SendUserStatsTxtDatabase: " & username & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub SendUserMiniStatsTxtFromDatabase(ByVal sendIndex As Integer, _
                                            ByVal username As String)

    '***************************************************
    'Author: Lorwik
    'Last Modification: 07/04/2021
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    If Not PersonajeExiste(username) Then
        Call WriteConsoleMsg(sendIndex, "Pj Inexistente", FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(sendIndex, "Estadisticas de: " & username, FontTypeNames.FONTTYPE_INFO)

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Connect
    #Else
        'Si perdimos la conexion reconectamos
        If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
    #End If
    
        query = "SELECT killed_npcs, killed_users, ciudadanos_matados, criminales_matados, class_id, genre_id, race_id FROM personaje WHERE UPPER(name) = (?)"
        
        If Not User_Database.MakeQuery(query, False, UCase$(username)) Then
            Call WriteConsoleMsg(sendIndex, "Pj Inexistente", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        Call WriteConsoleMsg(sendIndex, "Pj: " & username, FontTypeNames.FONTTYPE_INFO)
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
    Call LogDatabaseError("Error in SendUserMiniStatsTxtFromDatabase: " & username & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub SendUserOROTxtFromDatabase(ByVal sendIndex As Integer, _
                                      ByVal username As String)

    '***************************************************
    'Author: Lorwik
    'Last Modification: 07/04/2021
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query As String

    If Not PersonajeExiste(username) Then
        Call WriteConsoleMsg(sendIndex, "Pj Inexistente", FontTypeNames.FONTTYPE_INFO)
    Else
        #If DBConexionUnica = 0 Then
            Call User_Database.Database_Connect
        #Else
            'Si perdimos la conexion reconectamos
            If User_Database.CheckSQLStatus = False Then User_Database.Database_Reconnect
        #End If

        query = "SELECT bank_gold FROM personaje WHERE UPPER(name) = (?)"
        
        If Not User_Database.MakeQuery(query, False, UCase$(username)) Then
            Call WriteConsoleMsg(sendIndex, "Pj Inexistente", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        Call WriteConsoleMsg(sendIndex, "Pj: " & username, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(sendIndex, "Oro en banco: " & User_Database.Database_RecordSet!bank_gold, FontTypeNames.FONTTYPE_INFO)

        Set User_Database.Database_RecordSet = Nothing
        
        #If DBConexionUnica = 0 Then
            Call User_Database.Database_Close
        #End If

    End If

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in SendUserOROTxtFromDatabase: " & username & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub SendUserInvTxtFromDatabase(ByVal sendIndex As Integer, _
                                      ByVal username As String)

    '***************************************************
    'Author: Lorwik
    'Last Modification: 07/04/2021
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query   As String
    Dim LoopC   As Byte

    Dim ObjInd  As Long

    If Not PersonajeExiste(username) Then
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

        query = query & " FROM inventario_items WHERE user_id = (SELECT id FROM personaje WHERE UPPER(name) = (?))"
        
        If Not User_Database.MakeQuery(query, False, UCase$(username)) Then
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
    Call LogDatabaseError("Error in SendUserInvTxtFromDatabase: " & username & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub SendUserBovedaTxtFromDatabase(ByVal sendIndex As Integer, _
                                         ByVal username As String)

    '***************************************************
    'Author: Lorwik
    'Last Modification: 07/04/2021
    '***************************************************
    On Error GoTo ErrorHandler

    Dim query   As String
    Dim LoopC   As Byte

    Dim ObjInd As Long

    If Not PersonajeExiste(username) Then
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
        
        query = query & " FROM banco_items WHERE user_id = (SELECT id FROM personaje WHERE UPPER(name) = (?))"
        
        If Not User_Database.MakeQuery(query, False, UCase$(username)) Then
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
    Call LogDatabaseError("Error in SendUserBovedaTxtFromDatabase: " & username & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub SendCharacterInfoDatabase(ByVal UserIndex As Integer, ByVal username As String)

    '***************************************************
    'Author: Lorwik
    'Last Modification: 07/04/2021
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

    query = "SELECT race_id, class_id, genre_id, level, gold, bank_gold, rep_average, guild_requests_history, guild_index, guild_member_history, pertenece_real, pertenece_caos, ciudadanos_matados, criminales_matados FROM personaje WHERE UPPER(name) = (?)"
    
    If Not User_Database.MakeQuery(query, False, UCase$(username)) Then
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

    Call Protocol_Write.WriteCharacterInfo(UserIndex, username, User_Database.Database_RecordSet!race_id, User_Database.Database_RecordSet!class_id, User_Database.Database_RecordSet!genre_id, User_Database.Database_RecordSet!level, User_Database.Database_RecordSet!Gold, User_Database.Database_RecordSet!bank_gold, User_Database.Database_RecordSet!rep_average, SanitizeNullValue(User_Database.Database_RecordSet!guild_requests_history, vbNullString), gName, Miembro, User_Database.Database_RecordSet!pertenece_real, User_Database.Database_RecordSet!pertenece_caos, User_Database.Database_RecordSet!ciudadanos_matados, User_Database.Database_RecordSet!criminales_matados)

#If DBConexionUnica = 0 Then
    Call User_Database.Database_Close
#End If

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in SendCharacterInfoDatabase: " & username & ". " & Err.Number & " - " & Err.description)

End Sub

Public Function GetUserGuildMemberDatabase(ByVal username As String) As String

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

    query = "SELECT guild_member_history FROM personaje WHERE UPPER(name) = (?)"
    
    If Not User_Database.MakeQuery(query, False, UCase$(username)) Then
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
    Call LogDatabaseError("Error in GetUserGuildMemberDatabase: " & username & ". " & Err.Number & " - " & Err.description)

End Function

Public Function GetUserGuildAspirantDatabase(ByVal username As String) As Integer

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

    query = "SELECT guild_aspirant_index FROM personaje WHERE UPPER(name) = (?)"
    
    If Not User_Database.MakeQuery(query, False, UCase$(username)) Then
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
    Call LogDatabaseError("Error in GetUserGuildAspirantDatabase: " & username & ". " & Err.Number & " - " & Err.description)

End Function

Public Function GetUserGuildRejectionReasonDatabase(ByVal username As String) As String

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

    query = "SELECT guild_rejected_because FROM personaje WHERE UPPER(name) = (?)"
    
    If Not User_Database.MakeQuery(query, False, UCase$(username)) Then
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
    Call LogDatabaseError("Error in GetUserGuildRejectionReasonDatabase: " & username & ". " & Err.Number & " - " & Err.description)

End Function

Public Function GetUserGuildPedidosDatabase(ByVal username As String) As String

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

    query = "SELECT guild_requests_history FROM personaje WHERE UPPER(name) = (?)"
    
    If Not User_Database.MakeQuery(query, False, UCase$(username)) Then
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
    Call LogDatabaseError("Error in GetUserGuildPedidosDatabase: " & username & ". " & Err.Number & " - " & Err.description)

End Function

Public Sub SaveUserGuildRejectionReasonDatabase(ByVal username As String, _
                                                ByVal Reason As String)

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

    query = "UPDATE personaje SET guild_rejected_because = (?) WHERE UPPER(name) = (?)"
    Call User_Database.MakeQuery(query, True, Reason, UCase$(username))
    
#If DBConexionUnica = 0 Then
    Call User_Database.Database_Close
#End If

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in SaveUserGuildRejectionReasonDatabase: " & username & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub SaveUserGuildIndexDatabase(ByVal username As String, _
                                      ByVal GuildIndex As Integer)

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

    query = "UPDATE personaje SET guild_index = (?) WHERE UPPER(name) = (?)"
    Call User_Database.MakeQuery(query, True, GuildIndex, UCase$(username))
    
#If DBConexionUnica = 0 Then
    Call User_Database.Database_Close
#End If

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in SaveUserGuildIndexDatabase: " & username & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub SaveUserGuildAspirantDatabase(ByVal username As String, _
                                         ByVal AspirantIndex As Integer)

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

    query = "UPDATE personaje SET guild_aspirant_index = (?) WHERE UPPER(name) = (?)"
    Call User_Database.MakeQuery(query, True, AspirantIndex, UCase$(username))

#If DBConexionUnica = 0 Then
    Call User_Database.Database_Close
#End If

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in SaveUserGuildAspirantDatabase: " & username & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub SaveUserGuildMemberDatabase(ByVal username As String, ByVal guilds As String)

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

    query = "UPDATE personaje SET guild_member_history = (?) WHERE UPPER(name) = (?)"
    Call User_Database.MakeQuery(query, True, guilds, UCase$(username))

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
    #End If

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in SaveUserGuildMemberDatabase: " & username & ". " & Err.Number & " - " & Err.description)

End Sub

Public Sub SaveUserGuildPedidosDatabase(ByVal username As String, ByVal Pedidos As String)

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

    query = "UPDATE personaje SET guild_requests_history = (?) WHERE UPPER(name) = (?)"
    Call User_Database.MakeQuery(query, True, Pedidos, UCase$(username))

    #If DBConexionUnica = 0 Then
        Call User_Database.Database_Close
    #End If

    Exit Sub
ErrorHandler:
    Call LogDatabaseError("Error in SaveUserGuildPedidosDatabase: " & username & ". " & Err.Number & " - " & Err.description)

End Sub

Public Function SanitizeNullValue(ByVal Value As Variant, _
                                  ByVal defaultValue As Variant) As Variant
    SanitizeNullValue = IIf(IsNull(Value), defaultValue, Value)

End Function

