Attribute VB_Name = "Protocol_Handler"
'**************************************************************
' Protocol.bas - Handles all incoming / outgoing messages for client-server communications.
' Uses a binary protocol designed by myself.
'
' Designed and implemented by Juan Martin Sotuyo Dodero (Maraxus)
' (juansotuyo@gmail.com)
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
'Handles all incoming / outgoing packets for client - server communications
'The binary prtocol here used was designed by Juan Martin Sotuyo Dodero.
'This is the first time it's used in Alkon, though the second time it's coded.
'This implementation has several enhacements from the first design.
'
' @author Juan Martin Sotuyo Dodero (Maraxus) juansotuyo@gmail.com
' @version 1.0.0
' @date 20060517

Option Explicit

#If False Then

    Dim Map, X, Y, n, Mapa, race, helmet, weapon, shield, color, Value, errHandler, punishments, Length, obj, Index As Variant

#End If

''
'When we have a list of strings, we use this to separate them and prevent
'having too many string lengths in the queue. Yes, each string is NULL-terminated :P
Private Const SEPARATOR As String * 1 = vbNullChar

Private Enum ClientPacketID
    LoginExistingChar = 1           'OLOGIN
    LoginNewChar                    'NLOGIN
    Talk                            ';
    Yell                            '-
    Whisper                         '\
    Walk                            'M
    UseItem                         'USA
    RequestPositionUpdate           'RPU
    Attack                          'AT
    PickUp                          'AG
    SafeToggle                      '/SEG & SEG  (SEG's behaviour has to be coded in the client)
    CombatSafeToggle
    RequestGuildLeaderInfo          'GLINFO
    RequestAtributes                'ATR
    RequestFame                     'FAMA
    RequestSkills                   'ESKI
    RequestMiniStats                'FEST
    CommerceEnd                     'FINCOM
    UserCommerceEnd                 'FINCOMUSU
    UserCommerceConfirm
    CommerceChat
    BankEnd                         'FINBAN
    UserCommerceOk                  'COMUSUOK
    UserCommerceReject              'COMUSUNO
    Drop                            'TI
    CastSpell                       'LH
    LeftClick                       'LC
    AccionClick                     'RC
    Work                            'UK
    UseSpellMacro                   'UMH
    CraftearItem
    WorkClose
    WorkLeftClick                   'WLC
    InvitarPartyClick
    CreateNewGuild                  'CIG
    SpellInfo                      'INFS
    EquipItem                      'EQUI
    ChangeHeading                  'CHEA
    Train                          'ENTR
    CommerceBuy                    'COMP
    BankExtractItem                'RETI
    CommerceSell                   'VEND
    BankDeposit                    'DEPO
    ForumPost                      'DEMSG
    MoveSpell                      'DESPHE
    MoveBank
    ClanCodexUpdate               'DESCOD
    UserCommerceOffer             'OFRECER
    GuildAcceptPeace              'ACEPPEAT
    GuildRejectAlliance           'RECPALIA
    GuildRejectPeace              'RECPPEAT
    GuildAcceptAlliance           'ACEPALIA
    GuildOfferPeace               'PEACEOFF
    GuildOfferAlliance            'ALLIEOFF
    GuildAllianceDetails          'ALLIEDET
    GuildPeaceDetails             'PEACEDET
    GuildRequestJoinerInfo        'ENVCOMEN
    GuildAlliancePropList         'ENVALPRO
    GuildPeacePropList            'ENVPROPP
    GuildDeclareWar               'DECGUERR
    GuildNewWebsite               'NEWWEBSI
    GuildAcceptNewMember          'ACEPTARI
    GuildRejectNewMember          'RECHAZAR
    GuildKickMember               'ECHARCLA
    GuildUpdateNews               'ACTGNEWS
    GuildMemberInfo               '1HRINFO<
    GuildOpenElections            'ABREELEC
    GuildRequestMembership        'SOLICITUD
    GuildRequestDetails           'CLANDETAILS
    Online                        '/ONLINE
    Quit                          '/SALIR
    GuildLeave                    '/SALIRCLAN
    RequestAccountState           '/BALANCE
    PetStand                      '/QUIETO
    PetFollow                     '/ACOMPANAR
    ReleasePet                    '/LIBERAR
    TrainList                     '/ENTRENAR
    Rest                          '/DESCANSAR
    Meditate                      '/MEDITAR
    Resucitate                    '/RESUCITAR
    Heal                          '/CURAR
    Help                          '/AYUDA
    RequestStats                  '/EST
    CommerceStart                 '/COMERCIAR
    BankStart                     '/BOVEDA
    Enlist                        '/ENLISTAR
    Information                   '/INFORMACION
    Reward                        '/RECOMPENSA
    RequestMOTD                   '/MOTD
    UpTime                        '/UPTIME
    PartyLeave                    '/SALIRPARTY
    Inquiry                       '/ENCUESTA ( with no params )
    GuildMessage                  '/CMSG
    PartyMessage                  '/PMSG
    GuildOnline                   '/ONLINECLAN
    PartyOnline                   '/ONLINEPARTY
    CouncilMessage                '/BMSG
    RoleMasterRequest             '/ROL
    GMRequest                     '/GM
    ChangeDescription             '/DESC
    GuildVote                     '/VOTO
    punishments                   '/PENAS
    Gamble                        '/APOSTAR
    InquiryVote                   '/ENCUESTA ( with parameters )
    LeaveFaction                  '/RETIRAR ( with no arguments )
    BankExtractGold               '/RETIRAR ( with arguments )
    BankDepositGold               '/DEPOSITAR
    Denounce                      '/DENUNCIAR
    GuildFundate                  '/FUNDARCLAN
    GuildFundation
    PartyKick                     '/ECHARPARTY
    PartySetLeader                '/PARTYLIDER
    PartyAcceptMember             '/ACCEPTPARTY
    Ping                          '/PING
    RequestPartyForm
    Home
    ShowGuildNews
    ShareNpc                      '/COMPARTIR
    StopSharingNpc
    Consultation
    moveItem
    LoginExistingAccount      'CHOTS | Accounts
    Ecvc
    Acvc
    IrCvc
    DragAndDropHechizos
    Quest                       '/QUEST
    QuestAccept
    QuestListRequest
    QuestDetailsRequest
    QuestAbandon
    FightSend
    FightAccept
    CloseGuild
    Discord
    DeleteChar
    CraftsmanCreate
    AddAmigos
    DelAmigos
    OnAmigos
    MsgAmigos
    ChatGlobal
    AccionInventario
    invocar                     '/INVOCAR
    IniciarSubasta 'Iniciamos una subasta
    CancelarSubasta
    OfertarSubasta 'Ofertamos en la subasta
    ConsultaSubasta 'Si existe una subasta enviamos la Info, sino Abrimos el panel para iniciar una subasta.
    RespuestaInstruccion
    ShopInit
    BuyShop
    InitPVP
    DueloSet
    GMCommands
End Enum

Public Enum eGMCommands
    GMMessage = 1           '/GMSG
    showName                '/SHOWNAME
    OnlineRoyalArmy         '/ONLINEREAL
    OnlineChaosLegion       '/ONLINECAOS
    GoNearby                '/IRCERCA
    comment                 '/REM
    serverTime              '/HORA
    Where                   '/DONDE
    CreaturesInMap          '/NENE
    WarpMeToTarget          '/TELEPLOC
    WarpChar                '/TELEP
    Silence                 '/SILENCIAR
    SOSShowList             '/SHOW SOS
    SOSRemove               'SOSDONE
    GoToChar                '/IRA
    invisible               '/INVISIBLE
    GMPanel                 '/PANELGM
    RequestUserList         'LISTUSU
    Working                 '/TRABAJANDO
    Hiding                  '/OCULTANDO
    Jail                    '/CARCEL
    KillNPC                 '/RMATA
    WarnUser                '/ADVERTENCIA
    EditChar                '/MOD
    RequestCharInfo         '/INFO
    RequestCharStats        '/STAT
    RequestCharGold         '/BAL
    RequestCharInventory    '/INV
    RequestCharBank         '/BOV
    RequestCharSkills       '/SKILLS
    ReviveChar              '/REVIVIR
    OnlineGM                '/ONLINEGM
    OnlineMap               '/ONLINEMAP
    Forgive                 '/PERDON
    Kick                    '/ECHAR
    Execute                 '/EJECUTAR
    BanChar                 '/BAN
    UnbanChar               '/UNBAN
    NPCFollow               '/SEGUIR
    SummonChar              '/SUM
    SpawnListRequest        '/CC
    SpawnCreature           'SPA
    ResetNPCInventory       '/RESETINV
    ServerMessage           '/RMSG
    NickToIP                '/NICK2IP
    IPToNick                '/IP2NICK
    GuildOnlineMembers      '/ONCLAN
    TeleportCreate          '/CT
    TeleportDestroy         '/DT
    MeteoToggle             '/METEO
    SetCharDescription      '/SETDESC
    ForceMUSICToMap          '/FORCEMUSICMAP
    ForceWAVEToMap          '/FORCEWAVMAP
    RoyalArmyMessage        '/REALMSG
    ChaosLegionMessage      '/CAOSMSG
    CitizenMessage          '/CIUMSG
    CriminalMessage         '/CRIMSG
    TalkAsNPC               '/TALKAS
    DestroyAllItemsInArea   '/MASSDEST
    AcceptRoyalCouncilMember '/ACEPTCONSE
    AcceptChaosCouncilMember '/ACEPTCONSECAOS
    ItemsInTheFloor         '/PISO
    MakeDumb                '/ESTUPIDO
    MakeDumbNoMore          '/NOESTUPIDO
    DumpIPTables            '/DUMPSECURITY
    CouncilKick             '/KICKCONSE
    SetTrigger              '/TRIGGER
    AskTrigger              '/TRIGGER with no args
    BannedIPList            '/BANIPLIST
    BannedIPReload          '/BANIPRELOAD
    GuildMemberList         '/MIEMBROSCLAN
    GuildBan                '/BANCLAN
    BanIP                   '/BANIP
    UnbanIP                 '/UNBANIP
    CreateItem              '/CI
    DestroyItems            '/DEST
    ChaosLegionKick         '/NOCAOS
    RoyalArmyKick           '/NOREAL
    ForceMUSICAll           '/FORCEMUSIC
    ForceWAVEAll            '/FORCEWAV
    RemovePunishment        '/BORRARPENA
    TileBlockedToggle       '/BLOQ
    KillNPCNoRespawn        '/MATA
    KillAllNearbyNPCs       '/MASSKILL
    LastIP                  '/LASTIP
    ChangeMOTD              '/MOTDCAMBIA
    SetMOTD                 'ZMOTD
    SystemMessage           '/SMSG
    CreateNPC               '/ACC y /RACC
    ImperialArmour          '/AI1 - 4
    ChaosArmour             '/AC1 - 4
    NavigateToggle          '/NAVE
    ServerOpenToUsersToggle '/HABILITAR
    TurnOffServer           '/APAGAR
    TurnCriminal            '/CONDEN
    ResetFactions           '/RAJAR
    RemoveCharFromGuild     '/RAJARCLAN
    RequestCharMail         '/LASTEMAIL
    AlterName               '/ANAME
    DoBackUp                '/DOBACKUP
    ShowGuildMessages       '/SHOWCMSG
    SaveMap                 '/GUARDAMAPA
    ChangeZonaPK            '/MODZona PK
    ChangeZonaBackup        '/MODZona BACKUP
    ChangeZonaRestricted    '/MODZona RESTRINGIR
    ChangeZonaNoMagic       '/MODZona MAGIASINEFECTO
    ChangeZonaNoInvi        '/MODZona INVISINEFECTO
    ChangeZonaNoResu        '/MODZona RESUSINEFECTO
    ChangeZonaLand          '/MODZona TERRENO
    ChangeZonaZone          '/MODZona ZONA
    ChangeZonaStealNpc      '/MODZona ROBONPC
    ChangeZonaNoOcultar     '/MODZona OCULTARSINEFECTO
    ChangeZonaNoInvocar     '/MODZona INVOCARSINEFECTO
    SaveChars               '/GRABAR
    CleanSOS                '/BORRAR SOS
    ShowServerForm          '/SHOW INT
    night                   '/NOCHE
    KickAllChars            '/ECHARTODOSPJS
    ReloadNPCs              '/RELOADNPCS
    ReloadServerIni         '/RELOADSINI
    ReloadSpells            '/RELOADHECHIZOS
    ReloadObjects           '/RELOADOBJ
    Restart                 '/REINICIAR
    ResetAutoUpdate         '/AUTOUPDATE
    ChatColor               '/CHATCOLOR
    Ignored                 '/IGNORADO
    CheckSlot               '/SLOT
    SetIniVar               '/SETINIVAR LLAVE CLAVE VALOR
    CreatePretorianClan     '/CREARPRETORIANOS
    RemovePretorianClan     '/ELIMINARPRETORIANOS
    EnableDenounces         '/DENUNCIAS
    ShowDenouncesList       '/SHOW DENUNCIAS
    MapMessage              '/MAPMSG
    SetDialog               '/SETDIALOG
    Impersonate             '/IMPERSONAR
    Imitate                 '/MIMETIZAR
    RecordAdd
    RecordRemove
    RecordAddObs
    RecordListRequest
    RecordDetailsRequest
    ExitDestroy             '/DE
    SearchNpc               '/BUSCAR
    SearchObj               '/BUSCAR
    LimpiarMundo            '/LIMPIARMUNDO
    EditGems                '/EDITGEMS
    ConsultarGemas          '/CONSULTARGEMS
    SilenciarGlobal         '/SILENCIARGLOBAL
    ToggleGlobal            '/TOGGLEGLOBAL
    BanSerial
    UnBanSerial
    BanTemporal
End Enum

''
'The last existing client packet id.
Private Const LAST_CLIENT_PACKET_ID As Byte = 153

Public Enum FontTypeNames

    FONTTYPE_TALK
    FONTTYPE_FIGHT
    FONTTYPE_WARNING
    FONTTYPE_INFO
    FONTTYPE_INFOBOLD
    FONTTYPE_EJECUCION
    FONTTYPE_PARTY
    FONTTYPE_VENENO
    FONTTYPE_GUILD
    FONTTYPE_SERVER
    FONTTYPE_GUILDMSG
    FONTTYPE_CONSEJO
    FONTTYPE_CONSEJOCAOS
    FONTTYPE_CONSEJOVesA
    FONTTYPE_CONSEJOCAOSVesA
    FONTTYPE_CENTINELA
    FONTTYPE_GMMSG
    FONTTYPE_GM
    FONTTYPE_CITIZEN
    FONTTYPE_CONSE
    FONTTYPE_DIOS
    FONTTYPE_CRIMINAL
    FONTTYPE_EXP
    FONTTYPE_PRIVADO
    
End Enum

Public Enum eEditOptions

    eo_Gold = 1
    eo_Experience
    eo_Body
    eo_Head
    eo_CiticensKilled
    eo_CriminalsKilled
    eo_Level
    eo_Class
    eo_Skills
    eo_Nobleza
    eo_Asesino
    eo_Sex
    eo_Raza
    eo_addGold
    eo_Vida
    eo_Poss
    eo_Speed
    eo_ExperiencePVP

End Enum

''
' Handles incoming data.
'
' @param    userIndex The index of the user sending the message.

Public Function HandleIncomingData(ByVal UserIndex As Integer) As Boolean

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 01/09/07
    '
    '***************************************************
    On Error Resume Next
    
    With UserList(UserIndex)
        
        'Contamos cuantos paquetes recibimos.
        .Counters.PacketsTick = .Counters.PacketsTick + 1
        
        'Comento esto por ahora, por que cuando hago worldsave, envia mas paquetes en 40ms
        'y desconecta al pj, hay que reveer que hacer con esto y como solucionarlo.

        'Si recibis 10 paquetes en 40ms (intervalo del GameTimer), cierro la conexion.
        'If .Counters.PacketsTick > 10 Then
        '    Call CloseSocket(Userindex)
        '    Exit Function

        'End If

        'Se castea a long por que VB6 cuando usa SELECT CASE
        'Lo hace de manera mas efectiva https://www.gs-zone.org/temas/las-consecuencias-de-usar-byte-en-handleincomingdata.99245/
        Dim packetID As Long: packetID = CLng(.incomingData.PeekByte())
        
        'Debug.Print "packetID: " & packetID

        'Verifico si el paquete necesita que el user este logeado
        If Not (packetID = ClientPacketID.LoginExistingChar _
                Or packetID = ClientPacketID.LoginNewChar _
                Or packetID = ClientPacketID.LoginExistingAccount _
                Or packetID = ClientPacketID.DeleteChar) Then
             
            'Verifico si el user esta logeado
            If Not .flags.AccountLogged Then
                Call CloseSocket(UserIndex)
                Exit Function
                
            ElseIf Not .flags.UserLogged Then
                Call Cerrar_Usuario(UserIndex)
                Exit Function
                
            'El usuario ya logueo. Reseteamos el tiempo AFK si el ID es valido.
            ElseIf packetID <= LAST_CLIENT_PACKET_ID Then
                .Counters.IdleCount = 0
    
            End If
    
        ElseIf packetID <= LAST_CLIENT_PACKET_ID Then
            .Counters.IdleCount = 0
            
            'Vierifico si el user esta logeado
            If .flags.UserLogged Then
                Call CloseSocket(UserIndex)
                Exit Function
    
            End If
    
        End If
        
        ' Ante cualquier paquete, pierde la proteccion de ser atacado.
        .flags.NoPuedeSerAtacado = False
        
    End With
    
    Select Case packetID
        
        Case ClientPacketID.LoginExistingChar       'OLOGIN
            Call HandleLoginExistingChar(UserIndex)
        
        Case ClientPacketID.LoginNewChar            'NLOGIN
            Call HandleLoginNewChar(UserIndex)

        Case ClientPacketID.DeleteChar
            Call HandleDeleteChar(UserIndex)
        
        Case ClientPacketID.Talk                    ';
            Call HandleTalk(UserIndex)
        
        Case ClientPacketID.Yell                    '-
            Call HandleYell(UserIndex)
        
        Case ClientPacketID.Whisper                 '\
            Call HandleWhisper(UserIndex)
        
        Case ClientPacketID.Walk                    'M
            Call HandleWalk(UserIndex)
            
        Case ClientPacketID.UseItem                 'USA
            Call HandleUseItem(UserIndex)
        
        Case ClientPacketID.RequestPositionUpdate   'RPU
            Call HandleRequestPositionUpdate(UserIndex)
        
        Case ClientPacketID.Attack                  'AT
            Call HandleAttack(UserIndex)
        
        Case ClientPacketID.PickUp                  'AG
            Call HandlePickUp(UserIndex)
        
        Case ClientPacketID.SafeToggle              '/SEG & SEG  (SEG's behaviour has to be coded in the client)
            Call HandleSafeToggle(UserIndex)
        
        Case ClientPacketID.CombatSafeToggle
            Call HandleCombatToggle(UserIndex)
        
        Case ClientPacketID.RequestGuildLeaderInfo  'GLINFO
            Call HandleRequestGuildLeaderInfo(UserIndex)
        
        Case ClientPacketID.RequestAtributes        'ATR
            Call HandleRequestAtributes(UserIndex)
        
        Case ClientPacketID.RequestFame             'FAMA
            Call HandleRequestFame(UserIndex)
        
        Case ClientPacketID.RequestSkills           'ESKI
            Call HandleRequestSkills(UserIndex)
        
        Case ClientPacketID.RequestMiniStats        'FEST
            Call HandleRequestMiniStats(UserIndex)
        
        Case ClientPacketID.CommerceEnd             'FINCOM
            Call HandleCommerceEnd(UserIndex)
            
        Case ClientPacketID.CommerceChat
            Call HandleCommerceChat(UserIndex)
        
        Case ClientPacketID.UserCommerceEnd         'FINCOMUSU
            Call HandleUserCommerceEnd(UserIndex)
            
        Case ClientPacketID.UserCommerceConfirm
            Call HandleUserCommerceConfirm(UserIndex)
        
        Case ClientPacketID.BankEnd                 'FINBAN
            Call HandleBankEnd(UserIndex)
        
        Case ClientPacketID.UserCommerceOk          'COMUSUOK
            Call HandleUserCommerceOk(UserIndex)
        
        Case ClientPacketID.UserCommerceReject      'COMUSUNO
            Call HandleUserCommerceReject(UserIndex)
        
        Case ClientPacketID.Drop                    'TI
            Call HandleDrop(UserIndex)
        
        Case ClientPacketID.CastSpell               'LH
            Call HandleCastSpell(UserIndex)
        
        Case ClientPacketID.LeftClick               'LC
            Call HandleLeftClick(UserIndex)
        
        Case ClientPacketID.AccionClick             'RC
            Call HandleAccionClick(UserIndex)
        
        Case ClientPacketID.Work                    'UK
            Call HandleWork(UserIndex)
        
        Case ClientPacketID.UseSpellMacro           'UMH
            Call HandleUseSpellMacro(UserIndex)
        
        Case ClientPacketID.CraftearItem
            Call HandleCraftearItem(UserIndex)
            
        Case ClientPacketID.WorkClose
            Call HandleWorkClose(UserIndex)
        
        Case ClientPacketID.WorkLeftClick           'WLC
            Call HandleWorkLeftClick(UserIndex)
            
        Case ClientPacketID.InvitarPartyClick
            Call HandleInvitarPartyClick(UserIndex)
        
        Case ClientPacketID.CreateNewGuild          'CIG
            Call HandleCreateNewGuild(UserIndex)
            
        Case ClientPacketID.SpellInfo               'INFS
            Call HandleSpellInfo(UserIndex)
        
        Case ClientPacketID.EquipItem               'EQUI
            Call HandleEquipItem(UserIndex)
        
        Case ClientPacketID.ChangeHeading           'CHEA
            Call HandleChangeHeading(UserIndex)
        
        Case ClientPacketID.Train                   'ENTR
            Call HandleTrain(UserIndex)
        
        Case ClientPacketID.CommerceBuy             'COMP
            Call HandleCommerceBuy(UserIndex)
        
        Case ClientPacketID.BankExtractItem         'RETI
            Call HandleBankExtractItem(UserIndex)
        
        Case ClientPacketID.CommerceSell            'VEND
            Call HandleCommerceSell(UserIndex)
        
        Case ClientPacketID.BankDeposit             'DEPO
            Call HandleBankDeposit(UserIndex)
        
        Case ClientPacketID.ForumPost               'DEMSG
            Call HandleForumPost(UserIndex)
        
        Case ClientPacketID.MoveSpell               'DESPHE
            Call HandleMoveSpell(UserIndex)
            
        Case ClientPacketID.MoveBank
            Call HandleMoveBank(UserIndex)
        
        Case ClientPacketID.ClanCodexUpdate         'DESCOD
            Call HandleClanCodexUpdate(UserIndex)
        
        Case ClientPacketID.UserCommerceOffer       'OFRECER
            Call HandleUserCommerceOffer(UserIndex)
        
        Case ClientPacketID.GuildAcceptPeace        'ACEPPEAT
            Call HandleGuildAcceptPeace(UserIndex)
        
        Case ClientPacketID.GuildRejectAlliance     'RECPALIA
            Call HandleGuildRejectAlliance(UserIndex)
        
        Case ClientPacketID.GuildRejectPeace        'RECPPEAT
            Call HandleGuildRejectPeace(UserIndex)
        
        Case ClientPacketID.GuildAcceptAlliance     'ACEPALIA
            Call HandleGuildAcceptAlliance(UserIndex)
        
        Case ClientPacketID.GuildOfferPeace         'PEACEOFF
            Call HandleGuildOfferPeace(UserIndex)
        
        Case ClientPacketID.GuildOfferAlliance      'ALLIEOFF
            Call HandleGuildOfferAlliance(UserIndex)
        
        Case ClientPacketID.GuildAllianceDetails    'ALLIEDET
            Call HandleGuildAllianceDetails(UserIndex)
        
        Case ClientPacketID.GuildPeaceDetails       'PEACEDET
            Call HandleGuildPeaceDetails(UserIndex)
        
        Case ClientPacketID.GuildRequestJoinerInfo  'ENVCOMEN
            Call HandleGuildRequestJoinerInfo(UserIndex)
        
        Case ClientPacketID.GuildAlliancePropList   'ENVALPRO
            Call HandleGuildAlliancePropList(UserIndex)
        
        Case ClientPacketID.GuildPeacePropList      'ENVPROPP
            Call HandleGuildPeacePropList(UserIndex)
        
        Case ClientPacketID.GuildDeclareWar         'DECGUERR
            Call HandleGuildDeclareWar(UserIndex)
        
        Case ClientPacketID.GuildNewWebsite         'NEWWEBSI
            Call HandleGuildNewWebsite(UserIndex)
        
        Case ClientPacketID.GuildAcceptNewMember    'ACEPTARI
            Call HandleGuildAcceptNewMember(UserIndex)
        
        Case ClientPacketID.GuildRejectNewMember    'RECHAZAR
            Call HandleGuildRejectNewMember(UserIndex)
        
        Case ClientPacketID.GuildKickMember         'ECHARCLA
            Call HandleGuildKickMember(UserIndex)
        
        Case ClientPacketID.GuildUpdateNews         'ACTGNEWS
            Call HandleGuildUpdateNews(UserIndex)
        
        Case ClientPacketID.GuildMemberInfo         '1HRINFO<
            Call HandleGuildMemberInfo(UserIndex)
        
        Case ClientPacketID.GuildOpenElections      'ABREELEC
            Call HandleGuildOpenElections(UserIndex)
        
        Case ClientPacketID.GuildRequestMembership  'SOLICITUD
            Call HandleGuildRequestMembership(UserIndex)
        
        Case ClientPacketID.GuildRequestDetails     'CLANDETAILS
            Call HandleGuildRequestDetails(UserIndex)
                  
        Case ClientPacketID.Online                  '/ONLINE
            Call HandleOnline(UserIndex)
        
        Case ClientPacketID.Quit                    '/SALIR
            Call HandleQuit(UserIndex)
        
        Case ClientPacketID.GuildLeave              '/SALIRCLAN
            Call HandleGuildLeave(UserIndex)
        
        Case ClientPacketID.RequestAccountState     '/BALANCE
            Call HandleRequestAccountState(UserIndex)
        
        Case ClientPacketID.PetStand                '/QUIETO
            Call HandlePetStand(UserIndex)
        
        Case ClientPacketID.PetFollow               '/ACOMPANAR
            Call HandlePetFollow(UserIndex)
            
        Case ClientPacketID.ReleasePet              '/LIBERAR
            Call HandleReleasePet(UserIndex)
        
        Case ClientPacketID.TrainList               '/ENTRENAR
            Call HandleTrainList(UserIndex)
        
        Case ClientPacketID.Rest                    '/DESCANSAR
            Call HandleRest(UserIndex)
        
        Case ClientPacketID.Meditate                '/MEDITAR
            Call HandleMeditate(UserIndex)
        
        Case ClientPacketID.Resucitate              '/RESUCITAR
            Call HandleResucitate(UserIndex)
        
        Case ClientPacketID.Heal                    '/CURAR
            Call HandleHeal(UserIndex)
        
        Case ClientPacketID.Help                    '/AYUDA
            Call HandleHelp(UserIndex)
        
        Case ClientPacketID.RequestStats            '/EST
            Call HandleRequestStats(UserIndex)
        
        Case ClientPacketID.CommerceStart           '/COMERCIAR
            Call HandleCommerceStart(UserIndex)
        
        Case ClientPacketID.BankStart               '/BOVEDA
            Call HandleBankStart(UserIndex)
        
        Case ClientPacketID.Enlist                  '/ENLISTAR
            Call HandleEnlist(UserIndex)
        
        Case ClientPacketID.Information             '/INFORMACION
            Call HandleInformation(UserIndex)
        
        Case ClientPacketID.Reward                  '/RECOMPENSA
            Call HandleReward(UserIndex)
        
        Case ClientPacketID.RequestMOTD             '/MOTD
            Call HandleRequestMOTD(UserIndex)
        
        Case ClientPacketID.UpTime                  '/UPTIME
            Call HandleUpTime(UserIndex)
        
        Case ClientPacketID.PartyLeave              '/SALIRPARTY
            Call HandlePartyLeave(UserIndex)
        
        Case ClientPacketID.Inquiry                 '/ENCUESTA ( with no params )
            Call HandleInquiry(UserIndex)
        
        Case ClientPacketID.GuildMessage            '/CMSG
            Call HandleGuildMessage(UserIndex)
        
        Case ClientPacketID.PartyMessage            '/PMSG
            Call HandlePartyMessage(UserIndex)
        
        Case ClientPacketID.GuildOnline             '/ONLINECLAN
            Call HandleGuildOnline(UserIndex)
        
        Case ClientPacketID.PartyOnline             '/ONLINEPARTY
            Call HandlePartyOnline(UserIndex)
        
        Case ClientPacketID.CouncilMessage          '/BMSG
            Call HandleCouncilMessage(UserIndex)
        
        Case ClientPacketID.RoleMasterRequest       '/ROL
            Call HandleRoleMasterRequest(UserIndex)
        
        Case ClientPacketID.GMRequest               '/GM
            Call HandleGMRequest(UserIndex)
        
        Case ClientPacketID.ChangeDescription       '/DESC
            Call HandleChangeDescription(UserIndex)
        
        Case ClientPacketID.GuildVote               '/VOTO
            Call HandleGuildVote(UserIndex)
        
        Case ClientPacketID.punishments             '/PENAS
            Call HandlePunishments(UserIndex)
        
        Case ClientPacketID.Gamble                  '/APOSTAR
            Call HandleGamble(UserIndex)
        
        Case ClientPacketID.InquiryVote             '/ENCUESTA ( with parameters )
            Call HandleInquiryVote(UserIndex)
        
        Case ClientPacketID.LeaveFaction            '/RETIRAR ( with no arguments )
            Call HandleLeaveFaction(UserIndex)
        
        Case ClientPacketID.BankExtractGold         '/RETIRAR ( with arguments )
            Call HandleBankExtractGold(UserIndex)
        
        Case ClientPacketID.BankDepositGold         '/DEPOSITAR
            Call HandleBankDepositGold(UserIndex)
        
        Case ClientPacketID.Denounce                '/DENUNCIAR
            Call HandleDenounce(UserIndex)
        
        Case ClientPacketID.GuildFundate            '/FUNDARCLAN
            Call HandleGuildFundate(UserIndex)
            
        Case ClientPacketID.GuildFundation
            Call HandleGuildFundation(UserIndex)
        
        Case ClientPacketID.PartyKick               '/ECHARPARTY
            Call HandlePartyKick(UserIndex)
        
        Case ClientPacketID.PartySetLeader          '/PARTYLIDER
            Call HandlePartySetLeader(UserIndex)
        
        Case ClientPacketID.PartyAcceptMember       '/ACCEPTPARTY
            Call HandlePartyAcceptMember(UserIndex)
        
        Case ClientPacketID.Ping                    '/PING
            Call HandlePing(UserIndex)
            
        Case ClientPacketID.RequestPartyForm
            Call HandlePartyForm(UserIndex)
        
        Case ClientPacketID.Home
            Call HandleHome(UserIndex)
        
        Case ClientPacketID.ShowGuildNews
            Call HandleShowGuildNews(UserIndex)
            
        Case ClientPacketID.ShareNpc
            Call HandleShareNpc(UserIndex)
            
        Case ClientPacketID.StopSharingNpc
            Call HandleStopSharingNpc(UserIndex)
            
        Case ClientPacketID.Consultation
            Call HandleConsultation(UserIndex)
        
        Case ClientPacketID.moveItem
            Call HandleMoveItem(UserIndex)

        Case ClientPacketID.LoginExistingAccount
            Call HandleLoginExistingAccount(UserIndex)
            
        Case ClientPacketID.Ecvc
            Call HandleEnviaCvc(UserIndex)

        Case ClientPacketID.Acvc
            Call HandleAceptarCvc(UserIndex)

        Case ClientPacketID.IrCvc
            Call HandleIrCvc(UserIndex)
            
        Case ClientPacketID.DragAndDropHechizos
            Call HandleDragAndDropHechizos(UserIndex)
  
        Case ClientPacketID.Quest
            Call HandleQuest(UserIndex)
            
        Case ClientPacketID.QuestAccept
            Call HandleQuestAccept(UserIndex)
        
        Case ClientPacketID.QuestListRequest
            Call HandleQuestListRequest(UserIndex)
        
        Case ClientPacketID.QuestDetailsRequest
            Call HandleQuestDetailsRequest(UserIndex)
        
        Case ClientPacketID.QuestAbandon
            Call HandleQuestAbandon(UserIndex)

        Case ClientPacketID.FightSend
            Call HandleFightSend(UserIndex)
            
        Case ClientPacketID.FightAccept
            Call HandleFightAccept(UserIndex)
        
        Case ClientPacketID.CloseGuild
            Call HandleCloseGuild(UserIndex)
        
        Case ClientPacketID.Discord                  '/Discord
            Call HandleDiscord(UserIndex)
            
        Case ClientPacketID.CraftsmanCreate
            Call HandleCraftsmanCreate(UserIndex)
      
        Case ClientPacketID.AddAmigos
            Call HandleAddAmigo(UserIndex)

        Case ClientPacketID.DelAmigos
            Call HandleDelAmigo(UserIndex)

        Case ClientPacketID.OnAmigos
            Call HandleOnAmigo(UserIndex)

        Case ClientPacketID.MsgAmigos
            Call HandleMsgAmigo(UserIndex)
            
        Case ClientPacketID.ChatGlobal
            Call HandleChatGlobal(UserIndex)
            
        Case ClientPacketID.AccionInventario
            Call HandleAccionInventario(UserIndex)
            
        Case ClientPacketID.invocar
            Call HandleInvocar(UserIndex)
            
        Case ClientPacketID.IniciarSubasta
            Call HandleIniciaSubasta(UserIndex)
            
        Case ClientPacketID.CancelarSubasta
             Call HandleCancelarSubasta(UserIndex)
           
        Case ClientPacketID.OfertarSubasta
            Call HandleOfertaSubasta(UserIndex)
  
        Case ClientPacketID.ConsultaSubasta
            Call HandleConsultarSubasta(UserIndex)
            
        Case ClientPacketID.RespuestaInstruccion
            Call HandleRespuestaInstruccion(UserIndex)
            
        Case ClientPacketID.ShopInit
            Call HandleShopInit(UserIndex)
            
        Case ClientPacketID.BuyShop
            Call HandleBuyShop(UserIndex)
            
        Case ClientPacketID.InitPVP
            Call HandleInitPVP(UserIndex)
            
        Case ClientPacketID.DueloSet
            Call HandleDueloSet(UserIndex)
            
        Case ClientPacketID.GMCommands              'GM Messages
            Call HandleGMCommands(UserIndex)

        Case Else
            'ERROR : Abort!
            Call CloseSocket(UserIndex)

    End Select
    
    'Done with this packet, move on to next one or send everything if no more packets found
    If UserList(UserIndex).incomingData.Length > 0 And Err.Number = 0 Then
        Err.Clear
        HandleIncomingData = True
    
    ElseIf Err.Number <> 0 And Not Err.Number = UserList(UserIndex).incomingData.NotEnoughDataErrCode Then
        'An error ocurred, log it and kick player.
        Call LogError("Error: " & Err.Number & " [" & Err.description & "] " & " Source: " & Err.source & vbTab & " HelpFile: " & Err.HelpFile & vbTab & " HelpContext: " & Err.HelpContext & vbTab & " LastDllError: " & Err.LastDllError & vbTab & " - UserIndex: " & UserIndex & " - producido al manejar el paquete: " & CStr(packetID))
        Call CloseSocket(UserIndex)

        HandleIncomingData = False
    
    Else
        'Flush buffer - send everything that has been written
        Call FlushBuffer(UserIndex)

        HandleIncomingData = False

    End If

End Function

Private Sub HandleGMCommands(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo errHandler

    Dim Command As Byte

    With UserList(UserIndex)
        Call .incomingData.ReadByte
    
        Command = .incomingData.PeekByte
    
        Select Case Command

            Case eGMCommands.GMMessage                '/GMSG
                Call HandleGMMessage(UserIndex)
        
            Case eGMCommands.showName                '/SHOWNAME
                Call HandleShowName(UserIndex)
        
            Case eGMCommands.OnlineRoyalArmy
                Call HandleOnlineRoyalArmy(UserIndex)
        
            Case eGMCommands.OnlineChaosLegion       '/ONLINECAOS
                Call HandleOnlineChaosLegion(UserIndex)
        
            Case eGMCommands.GoNearby                '/IRCERCA
                Call HandleGoNearby(UserIndex)
        
            Case eGMCommands.comment                 '/REM
                Call HandleComment(UserIndex)
        
            Case eGMCommands.serverTime              '/HORA
                Call HandleServerTime(UserIndex)
        
            Case eGMCommands.Where                   '/DONDE
                Call HandleWhere(UserIndex)
        
            Case eGMCommands.CreaturesInMap          '/NENE
                Call HandleCreaturesInMap(UserIndex)
        
            Case eGMCommands.WarpMeToTarget          '/TELEPLOC
                Call HandleWarpMeToTarget(UserIndex)
        
            Case eGMCommands.WarpChar                '/TELEP
                Call HandleWarpChar(UserIndex)
        
            Case eGMCommands.Silence                 '/SILENCIAR
                Call HandleSilence(UserIndex)
        
            Case eGMCommands.SOSShowList             '/SHOW SOS
                Call HandleSOSShowList(UserIndex)
            
            Case eGMCommands.SOSRemove               'SOSDONE
                Call HandleSOSRemove(UserIndex)
        
            Case eGMCommands.GoToChar                '/IRA
                Call HandleGoToChar(UserIndex)
        
            Case eGMCommands.invisible               '/INVISIBLE
                Call HandleInvisible(UserIndex)
        
            Case eGMCommands.GMPanel                 '/PANELGM
                Call HandleGMPanel(UserIndex)
        
            Case eGMCommands.RequestUserList         'LISTUSU
                Call HandleRequestUserList(UserIndex)
        
            Case eGMCommands.Working                 '/TRABAJANDO
                Call HandleWorking(UserIndex)
        
            Case eGMCommands.Hiding                  '/OCULTANDO
                Call HandleHiding(UserIndex)
        
            Case eGMCommands.Jail                    '/CARCEL
                Call HandleJail(UserIndex)
        
            Case eGMCommands.KillNPC                 '/RMATA
                Call HandleKillNPC(UserIndex)
        
            Case eGMCommands.WarnUser                '/ADVERTENCIA
                Call HandleWarnUser(UserIndex)
        
            Case eGMCommands.EditChar                '/MOD
                Call HandleEditChar(UserIndex)
        
            Case eGMCommands.RequestCharInfo         '/INFO
                Call HandleRequestCharInfo(UserIndex)
        
            Case eGMCommands.RequestCharStats        '/STAT
                Call HandleRequestCharStats(UserIndex)
        
            Case eGMCommands.RequestCharGold         '/BAL
                Call HandleRequestCharGold(UserIndex)
        
            Case eGMCommands.RequestCharInventory    '/INV
                Call HandleRequestCharInventory(UserIndex)
        
            Case eGMCommands.RequestCharBank         '/BOV
                Call HandleRequestCharBank(UserIndex)
        
            Case eGMCommands.RequestCharSkills       '/SKILLS
                Call HandleRequestCharSkills(UserIndex)
        
            Case eGMCommands.ReviveChar              '/REVIVIR
                Call HandleReviveChar(UserIndex)
        
            Case eGMCommands.OnlineGM                '/ONLINEGM
                Call HandleOnlineGM(UserIndex)
        
            Case eGMCommands.OnlineMap               '/ONLINEMAP
                Call HandleOnlineMap(UserIndex)
        
            Case eGMCommands.Forgive                 '/PERDON
                Call HandleForgive(UserIndex)
        
            Case eGMCommands.Kick                    '/ECHAR
                Call HandleKick(UserIndex)
        
            Case eGMCommands.Execute                 '/EJECUTAR
                Call HandleExecute(UserIndex)
        
            Case eGMCommands.BanChar                 '/BAN
                Call HandleBanChar(UserIndex)
        
            Case eGMCommands.UnbanChar               '/UNBAN
                Call HandleUnbanChar(UserIndex)
        
            Case eGMCommands.NPCFollow               '/SEGUIR
                Call HandleNPCFollow(UserIndex)
        
            Case eGMCommands.SummonChar              '/SUM
                Call HandleSummonChar(UserIndex)
        
            Case eGMCommands.SpawnListRequest        '/CC
                Call HandleSpawnListRequest(UserIndex)
        
            Case eGMCommands.SpawnCreature           'SPA
                Call HandleSpawnCreature(UserIndex)
        
            Case eGMCommands.ResetNPCInventory       '/RESETINV
                Call HandleResetNPCInventory(UserIndex)
        
            Case eGMCommands.ServerMessage           '/RMSG
                Call HandleServerMessage(UserIndex)
        
            Case eGMCommands.MapMessage              '/MAPMSG
                Call HandleMapMessage(UserIndex)
            
            Case eGMCommands.NickToIP                '/NICK2IP
                Call HandleNickToIP(UserIndex)
        
            Case eGMCommands.IPToNick                '/IP2NICK
                Call HandleIPToNick(UserIndex)
        
            Case eGMCommands.GuildOnlineMembers      '/ONCLAN
                Call HandleGuildOnlineMembers(UserIndex)
        
            Case eGMCommands.TeleportCreate          '/CT
                Call HandleTeleportCreate(UserIndex)
        
            Case eGMCommands.TeleportDestroy         '/DT
                Call HandleTeleportDestroy(UserIndex)
        
            Case eGMCommands.MeteoToggle             '/METEO
                Call HandleMeteoToggle(UserIndex)
        
            Case eGMCommands.SetCharDescription      '/SETDESC
                Call HandleSetCharDescription(UserIndex)
        
            Case eGMCommands.ForceMUSICToMap          '/FORCEMUSICMAP
                Call HanldeForceMIDIToMap(UserIndex)
        
            Case eGMCommands.ForceWAVEToMap          '/FORCEWAVMAP
                Call HandleForceWAVEToMap(UserIndex)
        
            Case eGMCommands.RoyalArmyMessage        '/REALMSG
                Call HandleRoyalArmyMessage(UserIndex)
        
            Case eGMCommands.ChaosLegionMessage      '/CAOSMSG
                Call HandleChaosLegionMessage(UserIndex)
        
            Case eGMCommands.CitizenMessage          '/CIUMSG
                Call HandleCitizenMessage(UserIndex)
        
            Case eGMCommands.CriminalMessage         '/CRIMSG
                Call HandleCriminalMessage(UserIndex)
        
            Case eGMCommands.TalkAsNPC               '/TALKAS
                Call HandleTalkAsNPC(UserIndex)
        
            Case eGMCommands.DestroyAllItemsInArea   '/MASSDEST
                Call HandleDestroyAllItemsInArea(UserIndex)
        
            Case eGMCommands.AcceptRoyalCouncilMember '/ACEPTCONSE
                Call HandleAcceptRoyalCouncilMember(UserIndex)
        
            Case eGMCommands.AcceptChaosCouncilMember '/ACEPTCONSECAOS
                Call HandleAcceptChaosCouncilMember(UserIndex)
        
            Case eGMCommands.ItemsInTheFloor         '/PISO
                Call HandleItemsInTheFloor(UserIndex)
        
            Case eGMCommands.MakeDumb                '/ESTUPIDO
                Call HandleMakeDumb(UserIndex)
        
            Case eGMCommands.MakeDumbNoMore          '/NOESTUPIDO
                Call HandleMakeDumbNoMore(UserIndex)
        
            Case eGMCommands.DumpIPTables            '/DUMPSECURITY
                Call HandleDumpIPTables(UserIndex)
        
            Case eGMCommands.CouncilKick             '/KICKCONSE
                Call HandleCouncilKick(UserIndex)
        
            Case eGMCommands.SetTrigger              '/TRIGGER
                Call HandleSetTrigger(UserIndex)
        
            Case eGMCommands.AskTrigger              '/TRIGGER with no args
                Call HandleAskTrigger(UserIndex)
        
            Case eGMCommands.BannedIPList            '/BANIPLIST
                Call HandleBannedIPList(UserIndex)
        
            Case eGMCommands.BannedIPReload          '/BANIPRELOAD
                Call HandleBannedIPReload(UserIndex)
        
            Case eGMCommands.GuildMemberList         '/MIEMBROSCLAN
                Call HandleGuildMemberList(UserIndex)
        
            Case eGMCommands.GuildBan                '/BANCLAN
                Call HandleGuildBan(UserIndex)
        
            Case eGMCommands.BanIP                   '/BANIP
                Call HandleBanIP(UserIndex)
        
            Case eGMCommands.UnbanIP                 '/UNBANIP
                Call HandleUnbanIP(UserIndex)
        
            Case eGMCommands.CreateItem              '/CI
                Call HandleCreateItem(UserIndex)
        
            Case eGMCommands.DestroyItems            '/DEST
                Call HandleDestroyItems(UserIndex)
        
            Case eGMCommands.ChaosLegionKick         '/NOCAOS
                Call HandleChaosLegionKick(UserIndex)
        
            Case eGMCommands.RoyalArmyKick           '/NOREAL
                Call HandleRoyalArmyKick(UserIndex)
        
            Case eGMCommands.ForceMUSICAll            '/FORCEMUSIC
                Call HandleForceMUSICAll(UserIndex)
        
            Case eGMCommands.ForceWAVEAll            '/FORCEWAV
                Call HandleForceWAVEAll(UserIndex)
        
            Case eGMCommands.RemovePunishment        '/BORRARPENA
                Call HandleRemovePunishment(UserIndex)
        
            Case eGMCommands.TileBlockedToggle       '/BLOQ
                Call HandleTileBlockedToggle(UserIndex)
        
            Case eGMCommands.KillNPCNoRespawn        '/MATA
                Call HandleKillNPCNoRespawn(UserIndex)
        
            Case eGMCommands.KillAllNearbyNPCs       '/MASSKILL
                Call HandleKillAllNearbyNPCs(UserIndex)
        
            Case eGMCommands.LastIP                  '/LASTIP
                Call HandleLastIP(UserIndex)
        
            Case eGMCommands.ChangeMOTD              '/MOTDCAMBIA
                Call HandleChangeMOTD(UserIndex)
        
            Case eGMCommands.SetMOTD                 'ZMOTD
                Call HandleSetMOTD(UserIndex)
        
            Case eGMCommands.SystemMessage           '/SMSG
                Call HandleSystemMessage(UserIndex)
        
            Case eGMCommands.CreateNPC               '/ACC y /RACC
                Call HandleCreateNPC(UserIndex)
        
            Case eGMCommands.ImperialArmour          '/AI1 - 4
                Call HandleImperialArmour(UserIndex)
        
            Case eGMCommands.ChaosArmour             '/AC1 - 4
                Call HandleChaosArmour(UserIndex)
        
            Case eGMCommands.NavigateToggle          '/NAVE
                Call HandleNavigateToggle(UserIndex)
        
            Case eGMCommands.ServerOpenToUsersToggle '/HABILITAR
                Call HandleServerOpenToUsersToggle(UserIndex)
        
            Case eGMCommands.TurnOffServer           '/APAGAR
                Call HandleTurnOffServer(UserIndex)
        
            Case eGMCommands.TurnCriminal            '/CONDEN
                Call HandleTurnCriminal(UserIndex)
        
            Case eGMCommands.ResetFactions           '/RAJAR
                Call HandleResetFactions(UserIndex)
        
            Case eGMCommands.RemoveCharFromGuild     '/RAJARCLAN
                Call HandleRemoveCharFromGuild(UserIndex)
        
            Case eGMCommands.RequestCharMail         '/LASTEMAIL
                Call HandleRequestCharMail(UserIndex)
        
            Case eGMCommands.AlterName               '/ANAME
                Call HandleAlterName(UserIndex)
        
            Case eGMCommands.DoBackUp               '/DOBACKUP
                Call HandleDoBackUp(UserIndex)
        
            Case eGMCommands.ShowGuildMessages       '/SHOWCMSG
                Call HandleShowGuildMessages(UserIndex)
        
            Case eGMCommands.SaveMap                 '/GUARDAMAPA
                Call HandleSaveMap(UserIndex)
        
            Case eGMCommands.ChangeZonaPK         '/MODZona PK
                Call HandleChangeZonaPK(UserIndex)
            
            Case eGMCommands.ChangeZonaBackup     '/MODZona BACKUP
                Call HandleChangeZonaBackup(UserIndex)
        
            Case eGMCommands.ChangeZonaRestricted '/MODZona RESTRINGIR
                Call HandleChangeZonaRestricted(UserIndex)
        
            Case eGMCommands.ChangeZonaNoMagic    '/MODZona MAGIASINEFECTO
                Call HandleChangeZonaNoMagic(UserIndex)
        
            Case eGMCommands.ChangeZonaNoInvi     '/MODZona INVISINEFECTO
                Call HandleChangeZonaNoInvi(UserIndex)
        
            Case eGMCommands.ChangeZonaNoResu     '/MODZona RESUSINEFECTO
                Call HandleChangeZonaNoResu(UserIndex)
        
            Case eGMCommands.ChangeZonaLand       '/MODZona TERRENO
                Call HandleChangeZonaLand(UserIndex)
        
            Case eGMCommands.ChangeZonaZone       '/MODZona ZONA
                Call HandleChangeZonaZone(UserIndex)
        
            Case eGMCommands.ChangeZonaStealNpc   '/MODZona ROBONPC
                Call HandleChangeZonaStealNpc(UserIndex)
            
            Case eGMCommands.ChangeZonaNoOcultar  '/MODZona OCULTARSINEFECTO
                Call HandleChangeZonaNoOcultar(UserIndex)
            
            Case eGMCommands.ChangeZonaNoInvocar  '/MODZona INVOCARSINEFECTO
                Call HandleChangeZonaNoInvocar(UserIndex)
            
            Case eGMCommands.SaveChars               '/GRABAR
                Call HandleSaveChars(UserIndex)
        
            Case eGMCommands.CleanSOS                '/BORRAR SOS
                Call HandleCleanSOS(UserIndex)
        
            Case eGMCommands.ShowServerForm          '/SHOW INT
                Call HandleShowServerForm(UserIndex)
        
            Case eGMCommands.night                   '/NOCHE
                Call HandleNight(UserIndex)
        
            Case eGMCommands.KickAllChars            '/ECHARTODOSPJS
                Call HandleKickAllChars(UserIndex)
        
            Case eGMCommands.ReloadNPCs              '/RELOADNPCS
                Call HandleReloadNPCs(UserIndex)
        
            Case eGMCommands.ReloadServerIni         '/RELOADSINI
                Call HandleReloadServerIni(UserIndex)
        
            Case eGMCommands.ReloadSpells            '/RELOADHECHIZOS
                Call HandleReloadSpells(UserIndex)
        
            Case eGMCommands.ReloadObjects           '/RELOADOBJ
                Call HandleReloadObjects(UserIndex)
        
            Case eGMCommands.Restart                 '/REINICIAR
                Call HandleRestart(UserIndex)
        
            Case eGMCommands.ResetAutoUpdate         '/AUTOUPDATE
                Call HandleResetAutoUpdate(UserIndex)
        
            Case eGMCommands.ChatColor               '/CHATCOLOR
                Call HandleChatColor(UserIndex)
        
            Case eGMCommands.Ignored                 '/IGNORADO
                Call HandleIgnored(UserIndex)
        
            Case eGMCommands.CheckSlot               '/SLOT
                Call HandleCheckSlot(UserIndex)
        
            Case eGMCommands.SetIniVar               '/SETINIVAR LLAVE CLAVE VALOR
                Call HandleSetIniVar(UserIndex)
            
            Case eGMCommands.CreatePretorianClan     '/CREARPRETORIANOS
                Call HandleCreatePretorianClan(UserIndex)
         
            Case eGMCommands.RemovePretorianClan     '/ELIMINARPRETORIANOS
                Call HandleDeletePretorianClan(UserIndex)
                
            Case eGMCommands.EnableDenounces         '/DENUNCIAS
                Call HandleEnableDenounces(UserIndex)
            
            Case eGMCommands.ShowDenouncesList       '/SHOW DENUNCIAS
                Call HandleShowDenouncesList(UserIndex)
        
            Case eGMCommands.SetDialog               '/SETDIALOG
                Call HandleSetDialog(UserIndex)
            
            Case eGMCommands.Impersonate             '/IMPERSONAR
                Call HandleImpersonate(UserIndex)
            
            Case eGMCommands.Imitate                 '/MIMETIZAR
                Call HandleImitate(UserIndex)
            
            Case eGMCommands.RecordAdd
                Call HandleRecordAdd(UserIndex)
            
            Case eGMCommands.RecordAddObs
                Call HandleRecordAddObs(UserIndex)
            
            Case eGMCommands.RecordRemove
                Call HandleRecordRemove(UserIndex)
            
            Case eGMCommands.RecordListRequest
                Call HandleRecordListRequest(UserIndex)
            
            Case eGMCommands.RecordDetailsRequest
                Call HandleRecordDetailsRequest(UserIndex)
            
            Case eGMCommands.ExitDestroy
                Call HandleExitDestroy(UserIndex)
        
            Case eGMCommands.SearchNpc                          '/BUSCAR
                Call HandleSearchNpc(UserIndex)
           
            Case eGMCommands.SearchObj                          '/BUSCAR
                Call HandleSearchObj(UserIndex)
                                           
            Case eGMCommands.LimpiarMundo                       '/LIMPIARMUNDO
                Call HandleLimpiarMundo(UserIndex)
                
            Case eGMCommands.EditGems                           '/EDITGEMS
                Call HandleEditGems(UserIndex)
                
            Case eGMCommands.ConsultarGemas                     '/CONSULTARGEMS
                Call HandleConsultarGemas(UserIndex)
                
            Case eGMCommands.SilenciarGlobal
                Call HandleSilenciarGlobal(UserIndex)

            Case eGMCommands.ToggleGlobal
                Call HandleToggleGlobal(UserIndex)
                
            Case eGMCommands.BanSerial
                Call HandleBanSerial(UserIndex)
        
            Case eGMCommands.UnBanSerial
                Call HandleUnBanSerial(UserIndex)
                                           
            Case eGMCommands.BanTemporal
                Call HandleBanTemporal(UserIndex)
        End Select

    End With

    Exit Sub

errHandler:
    Call LogError("Error en GmCommands. Error: " & Err.Number & " - " & Err.description & ". Paquete: " & Command)

End Sub

Private Sub WriteConsoleServerUpTimeMsg(ByVal UserIndex As Integer)
    Dim time As Long
    Dim UpTimeStr As String
    
    'Get total time in seconds
    time = ((GetTickCount() And &H7FFFFFFF) - tInicioServer) \ 1000
    
    'Get times in dd:hh:mm:ss format
    UpTimeStr = (time Mod 60) & " segundos."
    time = time \ 60
    
    UpTimeStr = (time Mod 60) & " minutos, " & UpTimeStr
    time = time \ 60
    
    UpTimeStr = (time Mod 24) & " horas, " & UpTimeStr
    time = time \ 24
    
    If time = 1 Then
        UpTimeStr = time & " dia, " & UpTimeStr
    Else
        UpTimeStr = time & " dias, " & UpTimeStr
    End If

    Call WriteConsoleMsg(UserIndex, "Tiempo del Server Online: " & UpTimeStr, FontTypeNames.FONTTYPE_INFO)
End Sub


''
' Handles the "Home" message.
'
' @param    userIndex The index of the user sending the message.
Private Sub HandleHome(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Budi
    'Creation Date: 06/01/2010
    'Last Modification: 05/06/10
    'Pato - 05/06/10: Add the Ucase$ to prevent problems.
    '***************************************************
    With UserList(UserIndex)
        Call .incomingData.ReadByte

        If .flags.TargetNpcTipo = eNPCType.Gobernador Then
            
            If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 3 Then
                Call WriteConsoleMsg(UserIndex, "¡El gobernador no puede oirte, acercate mas para hablar con el!", FontTypeNames.FONTTYPE_INFO)
                
            Else
                Call setHome(UserIndex, Npclist(.flags.TargetNPC).Ciudad, .flags.TargetNPC)
                
            End If
        
        Else
            Call WriteConsoleMsg(UserIndex, "¡Debes seleccionar al gobernador de una ciudad para establecer un nuevo hogar!", FontTypeNames.FONTTYPE_INFO)
            
        End If

    End With

End Sub

''
' Handles the "DeleteChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDeleteChar(ByVal UserIndex As Integer)

'***************************************************
'Author: Lucas Recoaro (Recox)
'Last Modification: 07/01/20
'
'***************************************************
    If UserList(UserIndex).incomingData.Length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue
    Set Buffer = New clsByteQueue

    Call Buffer.CopyBuffer(UserList(UserIndex).incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte

    Dim PJSeleccionado As Byte
    PJSeleccionado = Buffer.ReadByte
    
    'If we got here then packet is complete, copy data back to original queue
    Call UserList(UserIndex).incomingData.CopyBuffer(Buffer)
    
    '¿Es un indice valido?
    If PJSeleccionado < 1 Or PJSeleccionado > MAXPJACCOUNTS Then
        Call WriteErrorMsg(UserIndex, "Error al borrar el PJ. Intentelo de nuevo o contacte con un Administrador.")
        Exit Sub
        
    End If
    
    If GetUserGuildIndexDatabase(UserList(UserIndex).AccountInfo.AccountPJ(PJSeleccionado).Name) > 0 Then
        Call WriteErrorMsg(UserIndex, "El personaje que intentas borrar pertenece a un clan. Debes salir del clan antes de borrar el personaje.")
        Exit Sub
        
    End If
    
    If NameIndex(UserList(UserIndex).AccountInfo.AccountPJ(PJSeleccionado).Name) > 0 Then
        Call WriteErrorMsg(UserIndex, "El personaje que intentas borrar esta conectado.")
        Exit Sub
        
    End If
    
    'Mandamos a borrar el PJ
    If BorrarUsuario(UserIndex, PJSeleccionado) Then
        'Si se pudo borrar enviamos paquete para mostrar mensaje satisfactorio en el cliente
        Call WriteDeletedChar(UserIndex)
        
    Else
        Call WriteErrorMsg(UserIndex, "Error al borrar el PJ. Intentelo de nuevo o contacte con un Administrador.")
        Exit Sub
        
    End If
    
    Exit Sub
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "LoginExistingChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleLoginExistingChar(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    Debug.Print UserList(UserIndex).incomingData.Length
    If UserList(UserIndex).incomingData.Length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue
    Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(UserList(UserIndex).incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte

    Dim SelectedID  As Byte
    Dim version     As String
    Dim username    As String
    
    SelectedID = Buffer.ReadByte
    
    'Convert version number to string
    version = CStr(Buffer.ReadByte()) & "." & CStr(Buffer.ReadByte()) & "." & CStr(Buffer.ReadByte())
    
    With UserList(UserIndex)
    
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

        'Obtenemos el nombre del PJ
        username = .AccountInfo.AccountPJ(SelectedID).Name
        
        If Not AsciiValidos(username) Then
            Call WriteErrorMsg(UserIndex, "Nombre invalido.")
            Call CloseUser(UserIndex)
            
            Exit Sub
    
        End If
        
        '¿El personaje existe?
        If Not PersonajeExiste(username) Then
            Call WriteErrorMsg(UserIndex, "El personaje no existe.")
            Call CloseUser(UserIndex)
            
            Exit Sub
    
        End If
    
        If BANCheck(username) Then
            Dim TiempoBan As Date
            
            TiempoBan = BanTimeCheck(username)
            
            If BanTimeCheck(username) Then
                '¿El ban expiro?
                If TiempoBan <= Now Then
                    Call UnBan(username)
                Else
                    Call WriteErrorMsg(UserIndex, "Se te ha prohibido la entrada a WinterAO hasta el " & Format(TiempoBan, "yyyy/mm/dd") & ". Puedes consultar el reglamento y el sistema de soporte desde http://winterao.com.ar")
                    Exit Sub
                End If
                
            Else
                Call WriteErrorMsg(UserIndex, "Se te ha prohibido la entrada a WinterAO debido a tu mal comportamiento. Puedes consultar el reglamento y el sistema de soporte desde http://winterao.com.ar")
                Exit Sub
                
            End If
            
        End If
        
        If Not VersionOK(version) Then
            Call WriteErrorMsg(UserIndex, "Esta version del juego es obsoleta, la version correcta es la " & ULTIMAVERSION & ". La misma se encuentra disponible en http://winterao.com.ar")
            
        Else
            Call ConnectUser(UserIndex, username)
            
        End If
    End With
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "LoginNewChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleLoginNewChar(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    Debug.Print UserList(UserIndex).incomingData.Length
    If UserList(UserIndex).incomingData.Length < 15 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue
    Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(UserList(UserIndex).incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte

    Dim username    As String
    Dim version     As String
    Dim race        As eRaza
    Dim gender      As eGenero
    Dim Class As eClass
    Dim Head As Integer
    Dim i As Byte
    
    username = Buffer.ReadASCIIString()

    'Convert version number to string
    version = CStr(Buffer.ReadByte()) & "." & CStr(Buffer.ReadByte()) & "." & CStr(Buffer.ReadByte())
    
    race = Buffer.ReadByte()
    gender = Buffer.ReadByte()
    Class = Buffer.ReadByte()
    Head = Buffer.ReadInteger
    
    'If we got here then packet is complete, copy data back to original queue
    Call UserList(UserIndex).incomingData.CopyBuffer(Buffer)
    
    If PuedeCrearPersonajes = 0 Then
        Call WriteErrorMsg(UserIndex, "La creacion de personajes en este servidor se ha deshabilitado.")
        Call CloseUser(UserIndex)
        Exit Sub
    End If
    
    If aClon.MaxPersonajes(UserList(UserIndex).IP) Then
        Call WriteErrorMsg(UserIndex, "Has creado demasiados personajes.")
        Call CloseUser(UserIndex)
        Exit Sub
    End If

    If GetCountUserAccount(UserIndex) >= 10 Then
        Call WriteErrorMsg(UserIndex, "No puedes crear mas de 10 personajes.")
        Call CloseUser(UserIndex)
        Exit Sub
    End If
                                        
    If Not VersionOK(version) Then
        Call WriteErrorMsg(UserIndex, "Esta version del juego es obsoleta, la version correcta es la " & ULTIMAVERSION & ". La misma se encuentra disponible en www.winterao.com.ar")
    Else
        Call ConnectNewUser(UserIndex, username, race, gender, Class, Head)
    End If
  
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "Talk" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTalk(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 13/01/2010
    '15/07/2009: ZaMa - Now invisible admins talk by console.
    '23/09/2009: ZaMa - Now invisible admins can't send empty chat.
    '13/01/2010: ZaMa - Now hidden on boat pirats recover the proper boat body.
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)
    
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim Chat As String
        
        Chat = Buffer.ReadASCIIString()
        
        '[Consejeros & GMs]
        If .flags.Privilegios And (PlayerType.Consejero Or PlayerType.SemiDios) Then
            Call LogGM(.Name, "Dijo: " & Chat)

        End If
        
        'I see you....
        If .flags.Oculto > 0 Then
            .flags.Oculto = 0
            .Counters.TiempoOculto = 0
            
            If .flags.Navegando = 1 Then
                If .clase = eClass.Pirat Then
                    ' Pierde la apariencia de fragata fantasmal
                    Call ToggleBoatBody(UserIndex)
                    Call WriteConsoleMsg(UserIndex, "Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
                    Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.Heading, NingunArma, NingunEscudo, NingunCasco, NingunAura, NingunAura)

                End If

            Else

                If .flags.invisible = 0 Then
                    Call UsUaRiOs.SetInvisible(UserIndex, UserList(UserIndex).Char.CharIndex, False)
                    Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If
        
        If LenB(Chat) <> 0 Then
            'Analize chat...
            Call Statistics.ParseChat(Chat)
            
            If Not (.flags.AdminInvisible = 1) Then
                If .flags.Muerto = 1 Then
                    Call SendData(SendTarget.ToDeadArea, UserIndex, PrepareMessageChatOverHead(Chat, .Char.CharIndex, 129, 129, 129))
                Else
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(Chat, .Char.CharIndex, .flags.ChatColor.r, .flags.ChatColor.g, .flags.ChatColor.b))

                End If

            Else

                If RTrim(Chat) <> "" Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageConsoleMsg("Gm> " & Chat, FontTypeNames.FONTTYPE_GM))

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "Yell" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleYell(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 13/01/2010 (ZaMa)
    '15/07/2009: ZaMa - Now invisible admins yell by console.
    '13/01/2010: ZaMa - Now hidden on boat pirats recover the proper boat body.
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)
    
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim Chat As String
        
        Chat = Buffer.ReadASCIIString()

        '[Consejeros & GMs]
        If .flags.Privilegios And (PlayerType.Consejero Or PlayerType.SemiDios) Then
            Call LogGM(.Name, "Grito: " & Chat)

        End If
            
        'I see you....
        If .flags.Oculto > 0 Then
            .flags.Oculto = 0
            .Counters.TiempoOculto = 0
            
            If .flags.Navegando = 1 Then
                If .clase = eClass.Pirat Then
                    ' Pierde la apariencia de fragata fantasmal
                    Call ToggleBoatBody(UserIndex)
                    Call WriteConsoleMsg(UserIndex, "Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
                    Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.Heading, NingunArma, NingunEscudo, NingunCasco, NingunAura, NingunAura)

                End If

            Else

                If .flags.invisible = 0 Then
                    Call UsUaRiOs.SetInvisible(UserIndex, .Char.CharIndex, False)
                    Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If
            
        If LenB(Chat) <> 0 Then
            'Analize chat...
            Call Statistics.ParseChat(Chat)
                
            If .flags.Privilegios And PlayerType.User Then
                If UserList(UserIndex).flags.Muerto = 1 Then
                    Call SendData(SendTarget.ToDeadArea, UserIndex, PrepareMessageChatOverHead(Chat, .Char.CharIndex, 129, 129, 129))
                Else
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(Chat, .Char.CharIndex, 255, 0, 0))

                End If

            Else

                If Not (.flags.AdminInvisible = 1) Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead(Chat, .Char.CharIndex, 255, 255, 0))
                Else
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageConsoleMsg("Gm> " & Chat, FontTypeNames.FONTTYPE_GM))

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "Whisper" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWhisper(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 03/12/2010
    '28/05/2009: ZaMa - Now it doesn't appear any message when private talking to an invisible admin
    '15/07/2009: ZaMa - Now invisible admins wisper by console.
    '03/12/2010: Enanoh - Agregue susurro a Admins en modo consulta y Los Dioses pueden susurrar en ciertos casos.
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim Chat            As String

        Dim TargetUserIndex As Integer

        Dim TargetPriv      As PlayerType

        Dim userPriv        As PlayerType

        Dim TargetName      As String
        
        TargetName = Buffer.ReadASCIIString()
        Chat = Buffer.ReadASCIIString()
        
        userPriv = .flags.Privilegios
        
        If .flags.Muerto Then
            Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
            'Call WriteConsoleMsg(UserIndex, "Estas muerto!! Los muertos no pueden comunicarse con el mundo de los vivos. ", FontTypeNames.FONTTYPE_INFO)
        Else
            ' Offline?
            TargetUserIndex = NameIndex(TargetName)

            If TargetUserIndex = INVALID_INDEX Then

                ' Admin?
                If EsGmChar(TargetName) Then
                    Call WriteConsoleMsg(UserIndex, "No puedes susurrarle a los Administradores.", FontTypeNames.FONTTYPE_INFO)
                    ' Whisperer admin? (Else say nothing)
                ElseIf (userPriv And (PlayerType.Dios Or PlayerType.Admin)) <> 0 Then
                    Call WriteConsoleMsg(UserIndex, "Usuario inexistente.", FontTypeNames.FONTTYPE_INFO)

                End If
                
                ' Online
            Else
                ' Privilegios
                TargetPriv = UserList(TargetUserIndex).flags.Privilegios
                
                ' Consejeros, semis y usuarios no pueden susurrar a dioses (Salvo en consulta)
                If (TargetPriv And (PlayerType.Dios Or PlayerType.Admin)) <> 0 And (userPriv And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios)) <> 0 And Not .flags.EnConsulta Then
                    
                    ' No puede
                    Call WriteConsoleMsg(UserIndex, "No puedes susurrarle a los Administradores.", FontTypeNames.FONTTYPE_INFO)

                    ' Usuarios no pueden susurrar a semis o conses (Salvo en consulta)
                ElseIf (userPriv And PlayerType.User) <> 0 And (Not TargetPriv And PlayerType.User) <> 0 And Not .flags.EnConsulta Then
                    
                    ' No puede
                    Call WriteConsoleMsg(UserIndex, "No puedes susurrarle a los Administradores.", FontTypeNames.FONTTYPE_INFO)

                Else

                    '[Consejeros & GMs]
                    If userPriv And (PlayerType.Consejero Or PlayerType.SemiDios) Then
                        Call LogGM(.Name, "Le susurro a '" & UserList(TargetUserIndex).Name & "' " & Chat)
                    
                        ' Usuarios a administradores
                    ElseIf (userPriv And PlayerType.User) <> 0 And (TargetPriv And PlayerType.User) = 0 Then
                        Call LogGM(UserList(TargetUserIndex).Name, .Name & " le susurro en consulta: " & Chat)

                    End If
                    
                    If LenB(Chat) <> 0 Then
                        'Analize chat...
                        Call Statistics.ParseChat(Chat)
                        
                        ' Dios susurrando a distancia
                        If Not EstaPCarea(UserIndex, TargetUserIndex) Then
                            
                            Call WriteConsoleMsg(UserIndex, UserList(UserIndex).Name & "> " & Chat, FontTypeNames.FONTTYPE_PRIVADO)
                            Call WriteConsoleMsg(TargetUserIndex, UserList(UserIndex).Name & "> " & Chat, FontTypeNames.FONTTYPE_PRIVADO)
                            
                        ElseIf Not (.flags.AdminInvisible = 1) Then
                            Call WriteChatOverHead(UserIndex, Chat, .Char.CharIndex, 0, 192, 0, True)
                            Call WriteChatOverHead(TargetUserIndex, Chat, .Char.CharIndex, 0, 192, 0, True)
                            Call WriteConsoleMsg(UserIndex, UserList(UserIndex).Name & "> " & Chat, FontTypeNames.FONTTYPE_PRIVADO)
                            Call WriteConsoleMsg(TargetUserIndex, UserList(UserIndex).Name & "> " & Chat, FontTypeNames.FONTTYPE_PRIVADO)

                        Else
                            Call WriteConsoleMsg(UserIndex, UserList(UserIndex).Name & "> " & Chat, FontTypeNames.FONTTYPE_PRIVADO)

                            If UserIndex <> TargetUserIndex Then Call WriteConsoleMsg(TargetUserIndex, UserList(TargetUserIndex).Name & "> " & Chat, FontTypeNames.FONTTYPE_PRIVADO)
                            

                        End If

                    End If

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "Walk" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWalk(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 15/09/2024 (Lorwik)
    '11/19/09 Pato - Now the class bandit can walk hidden.
    '13/01/2010: ZaMa - Now hidden on boat pirats recover the proper boat body.
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    Dim dummy    As Long

    Dim TempTick As Long

    Dim Heading  As eHeading
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Heading = .incomingData.ReadByte()
        
        '¿Está trabajando?
        If .flags.MacroTrabajo <> 0 Then
            Call DejardeTrabajar(UserIndex)
        End If
        
        If .flags.Paralizado = 0 Or .flags.Inmovilizado = 0 Then
        
            If .flags.Meditando Then
                'Stop meditating, next action will start movement.
                .flags.Meditando = False
                .Char.FX = 0
                .Char.loops = 0

                Call WriteConsoleMsg(UserIndex, "Dejas de meditar.", FontTypeNames.FONTTYPE_INFO)
                
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, 0, 0))
            End If
        
            Dim CurrentTick As Long
            CurrentTick = GetTickCount
            
            'Prevent SpeedHack (refactored by WyroX)
            If Not EsGm(UserIndex) And .flags.Velocidad > 0 Then
                Dim ElapsedTimeStep As Long, MinTimeStep As Long, DeltaStep As Single
                ElapsedTimeStep = CurrentTick - .Counters.LastStep
                MinTimeStep = IntervaloCaminar / .flags.Velocidad
                DeltaStep = (MinTimeStep - ElapsedTimeStep) / MinTimeStep

                If DeltaStep > 0 Then
                
                    .Counters.SpeedHackCounter = .Counters.SpeedHackCounter + DeltaStep
                
                    If .Counters.SpeedHackCounter > MaximoSpeedHack Then
                        'Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Administración » Posible uso de SpeedHack del usuario " & .name & ".", e_FontTypeNames.FONTTYPE_SERVER))
                        Call WritePosUpdate(UserIndex)
                        Exit Sub

                    End If

                Else
                
                    .Counters.SpeedHackCounter = .Counters.SpeedHackCounter + DeltaStep * 5

                    If .Counters.SpeedHackCounter < 0 Then .Counters.SpeedHackCounter = 0

                End If

            End If
        
            'Move user
            If MoveUserChar(UserIndex, Heading) Then
            
                'If exiting, cancel
                Call CancelExit(UserIndex)
        
                'Si esta casteando, lo cancelamos
                Call CancelCast(UserIndex)
                
                'Stop resting if needed
                If .flags.Descansar Then
                    .flags.Descansar = False
                    
                    Call WriteRestOK(UserIndex)
                    Call WriteConsoleMsg(UserIndex, "Has dejado de descansar.", FontTypeNames.FONTTYPE_INFO)

                End If
                
            Else
                .Counters.LastStep = 0
                Call WritePosUpdate(UserIndex)
            
            End If

        Else    'paralized

            If Not .flags.UltimoMensaje = 1 Then
                .flags.UltimoMensaje = 1
                
                Call WriteConsoleMsg(UserIndex, "No puedes moverte porque estas paralizado.", FontTypeNames.FONTTYPE_INFO)

            End If

        End If
        
        'Can't move while hidden except he is a thief
        If .flags.Oculto = 1 And .flags.AdminInvisible = 0 Then
            If .clase <> eClass.Thief And .clase <> eClass.Bandit Then
                .flags.Oculto = 0
                .Counters.TiempoOculto = 0
            
                If .flags.Navegando = 1 Then
                    If .clase = eClass.Pirat Then
                        ' Pierde la apariencia de fragata fantasmal
                        Call ToggleBoatBody(UserIndex)
                        Call WriteConsoleMsg(UserIndex, "Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
                        Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.Heading, NingunArma, NingunEscudo, NingunCasco, NingunAura, NingunAura)

                    End If

                Else

                    'If not under a spell effect, show char
                    If .flags.invisible = 0 Then
                        Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
                        Call UsUaRiOs.SetInvisible(UserIndex, .Char.CharIndex, False)

                    End If

                End If

            End If

        End If

    End With

End Sub

''
' Handles the "RequestPositionUpdate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestPositionUpdate(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    UserList(UserIndex).incomingData.ReadByte
    
    Call WritePosUpdate(UserIndex)

End Sub

''
' Handles the "Attack" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleAttack(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 13/01/2010
    'Last Modified By: ZaMa
    '10/01/2008: Tavo - Se cancela la salida del juego si el user esta saliendo.
    '13/11/2009: ZaMa - Se cancela el estado no atacable al atcar.
    '13/01/2010: ZaMa - Now hidden on boat pirats recover the proper boat body.
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'If dead, can't attack
        If .flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "Estas muerto!!", FontTypeNames.FONTTYPE_INFO)
            Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
            Exit Sub

        End If
        
        'If user meditates, can't attack
        If .flags.Meditando Then
            Exit Sub

        End If
        
        '¿Está trabajando?
        If .flags.MacroTrabajo <> 0 Then
            Call WriteConsoleMsg(UserIndex, "¡Estas trabajando!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'If equiped weapon is ranged, can't attack this way
        If .Invent.WeaponEqpObjIndex > 0 Then
            If ObjData(.Invent.WeaponEqpObjIndex).proyectil = 1 Then
                Call WriteConsoleMsg(UserIndex, "No puedes usar asi este arma.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

        End If
        
        'If exiting, cancel
        Call CancelExit(UserIndex)
        
        'Si esta casteando, lo cancelamos
        Call CancelCast(UserIndex)
        
        'Play AttackAnim on Clients
        Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageCharacterAttackAnim(.Char.CharIndex))
        
        'Attack!
        Call UsuarioAtaca(UserIndex)
        
        'Now you can be atacked
        .flags.NoPuedeSerAtacado = False
        
        'I see you...
        If .flags.Oculto > 0 And .flags.AdminInvisible = 0 Then
            .flags.Oculto = 0
            .Counters.TiempoOculto = 0
            
            If .flags.Navegando = 1 Then
                If .clase = eClass.Pirat Then
                    ' Pierde la apariencia de fragata fantasmal
                    Call ToggleBoatBody(UserIndex)
                    Call WriteConsoleMsg(UserIndex, "Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
                    Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.Heading, NingunArma, NingunEscudo, NingunCasco, NingunAura, NingunAura)

                End If

            Else

                If .flags.invisible = 0 Then
                    Call UsUaRiOs.SetInvisible(UserIndex, .Char.CharIndex, False)
                    Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If

    End With

End Sub

''
' Handles the "PickUp" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePickUp(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 07/25/09
    '02/26/2006: Marco - Agregue un checkeo por si el usuario trata de agarrar un item mientras comercia.
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'If dead, it can't pick up objects
        If .flags.Muerto = 1 Then Exit Sub
        
        'If user is trading items and attempts to pickup an item, he's cheating, so we kick him.
        If .flags.Comerciando Then Exit Sub
        
        'Lower rank administrators can't pick up items
        If .flags.Privilegios And PlayerType.Consejero Then
            If Not .flags.Privilegios And PlayerType.RoleMaster Then
                Call WriteConsoleMsg(UserIndex, "No puedes tomar ningUn objeto.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

        End If
        
        Call GetObj(UserIndex)

    End With

End Sub

''
' Handles the "SafeToggle" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSafeToggle(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Seguro Then
            Call WriteMultiMessage(UserIndex, eMessages.SafeModeOff) 'Call WriteSafeModeOff(UserIndex)
        Else
            Call WriteMultiMessage(UserIndex, eMessages.SafeModeOn) 'Call WriteSafeModeOn(UserIndex)

        End If
        
        .flags.Seguro = Not .flags.Seguro

    End With

End Sub

''
' Handles the "CombatSafeToggle" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCombatToggle(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Lorwik
    'Creation Date: 23/10/2020
    '***************************************************
    With UserList(UserIndex)
        Call .incomingData.ReadByte
        
        .flags.ModoCombate = Not .flags.ModoCombate
        
        If .flags.ModoCombate Then
            Call WriteMultiMessage(UserIndex, eMessages.CombatSafeOn) 'Call WriteCombatSafeOn(UserIndex)
        Else
            Call WriteMultiMessage(UserIndex, eMessages.CombatSafeOff) 'Call WriteCombatSafeOff(UserIndex)

        End If

    End With

End Sub

''
' Handles the "RequestGuildLeaderInfo" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestGuildLeaderInfo(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    UserList(UserIndex).incomingData.ReadByte
    
    Call modGuilds.SendGuildLeaderInfo(UserIndex)

End Sub

''
' Handles the "RequestAtributes" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestAtributes(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    Call WriteAttributes(UserIndex)

End Sub

''
' Handles the "RequestFame" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestFame(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    Call EnviarFama(UserIndex)

End Sub

''
' Handles the "RequestSkills" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestSkills(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    Call WriteSendSkills(UserIndex)

End Sub

''
' Handles the "RequestMiniStats" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestMiniStats(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    Call WriteMiniStats(UserIndex)

End Sub

''
' Handles the "CommerceEnd" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCommerceEnd(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    'User quits commerce mode
    UserList(UserIndex).flags.Comerciando = False
    Call WriteCommerceEnd(UserIndex)

End Sub

''
' Handles the "UserCommerceEnd" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUserCommerceEnd(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 11/03/2010
    '11/03/2010: ZaMa - Le avisa por consola al que cencela que dejo de comerciar.
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Quits commerce mode with user
        If .ComUsu.DestUsu > 0 Then
            If UserList(.ComUsu.DestUsu).ComUsu.DestUsu = UserIndex Then
                Call WriteConsoleMsg(.ComUsu.DestUsu, .Name & " ha dejado de comerciar con vos.", FontTypeNames.FONTTYPE_TALK)
                Call FinComerciarUsu(.ComUsu.DestUsu)

            End If

        End If
        
        Call FinComerciarUsu(UserIndex)
        Call WriteConsoleMsg(UserIndex, "Has dejado de comerciar.", FontTypeNames.FONTTYPE_TALK)

    End With

End Sub

''
' Handles the "UserCommerceConfirm" message.
'
' @param    userIndex The index of the user sending the message.
Private Sub HandleUserCommerceConfirm(ByVal UserIndex As Integer)
    '***************************************************
    'Author: ZaMa
    'Last Modification: 14/12/2009
    '
    '***************************************************
    
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte

    'Validate the commerce
    If PuedeSeguirComerciando(UserIndex) Then
        'Tell the other user the confirmation of the offer
        Call WriteUserOfferConfirm(UserList(UserIndex).ComUsu.DestUsu)
        UserList(UserIndex).ComUsu.Confirmo = True

    End If
    
End Sub

Private Sub HandleCommerceChat(ByVal UserIndex As Integer)

    '***************************************************
    'Author: ZaMa
    'Last Modification: 03/12/2009
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)
    
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim Chat As String
        
        Chat = Buffer.ReadASCIIString()
        
        If LenB(Chat) <> 0 Then
            If PuedeSeguirComerciando(UserIndex) Then
                'Analize chat...
                Call Statistics.ParseChat(Chat)
                
                Chat = UserList(UserIndex).Name & "> " & Chat
                Call WriteCommerceChat(UserIndex, Chat, FontTypeNames.FONTTYPE_PARTY)
                Call WriteCommerceChat(UserList(UserIndex).ComUsu.DestUsu, Chat, FontTypeNames.FONTTYPE_PARTY)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "BankEnd" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankEnd(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'User exits banking mode
        .flags.Comerciando = False
        Call WriteBankEnd(UserIndex)

    End With

End Sub

''
' Handles the "UserCommerceOk" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUserCommerceOk(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    'Trade accepted
    Call AceptarComercioUsu(UserIndex)

End Sub

''
' Handles the "UserCommerceReject" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUserCommerceReject(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    Dim otherUser As Integer
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        otherUser = .ComUsu.DestUsu
        
        'Offer rejected
        If otherUser > 0 Then
            If UserList(otherUser).flags.UserLogged Then
                Call WriteConsoleMsg(otherUser, .Name & " ha rechazado tu oferta.", FontTypeNames.FONTTYPE_TALK)
                Call FinComerciarUsu(otherUser)

            End If

        End If
        
        Call WriteConsoleMsg(UserIndex, "Has rechazado la oferta del otro usuario.", FontTypeNames.FONTTYPE_TALK)
        Call FinComerciarUsu(UserIndex)

    End With

End Sub

''
' Handles the "Drop" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDrop(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 07/25/09
    '07/25/09: Marco - Agregue un checkeo para patear a los usuarios que tiran items mientras comercian.
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    Dim Slot   As Byte

    Dim Amount As Integer
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Slot = .incomingData.ReadByte()
        Amount = .incomingData.ReadInteger()

        'low rank admins can't drop item. Neither can the dead nor those sailing.
        If .flags.Navegando = 1 Or .flags.Muerto = 1 Or ((.flags.Privilegios And PlayerType.Consejero) <> 0 And (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0) Then Exit Sub

        '¿Está trabajando?
        If .flags.MacroTrabajo <> 0 Then
            Call WriteConsoleMsg(UserIndex, "¡Estas trabajando!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        '¿Puede tirar items en el mapa?
        If MapZonas(.Pos.Map, UserZonaId(UserIndex)).NoTirarItems = True Then
            Call WriteConsoleMsg(UserIndex, "No puedes tirar objetos en el mapa.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        'If the user is trading, he can't drop items => He's cheating, we kick him.
        If .flags.Comerciando Then Exit Sub

        'Are we dropping gold or other items??
        If Slot = FLAGORO Then
            If Amount > 10000 Then Exit Sub 'Don't drop too much gold

            Call TirarOro(Amount, UserIndex)
            
            Call WriteUpdateGold(UserIndex)
        Else

            'Only drop valid slots
            If Slot <= MAX_INVENTORY_SLOTS And Slot > 0 Then
                If .Invent.Object(Slot).ObjIndex = 0 Then
                    Exit Sub

                End If
                
                Call DropObj(UserIndex, Slot, Amount, .Pos.Map, .Pos.X, .Pos.Y)
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_DROP, .Pos.X, .Pos.Y))

            End If

        End If

    End With

End Sub

''
' Handles the "CastSpell" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCastSpell(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '13/11/2009: ZaMa - Ahora los npcs pueden atacar al usuario si quizo castear un hechizo
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Spell As Byte
        
        Spell = .incomingData.ReadByte()
        
        If .flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "Estas muerto!!", FontTypeNames.FONTTYPE_INFO)
            Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
            Exit Sub

        End If
        
        '¿Está trabajando?
        If .flags.MacroTrabajo <> 0 Then
            Call WriteConsoleMsg(UserIndex, "¡Estas trabajando!", FontTypeNames.FONTTYPE_INFOBOLD)
            Exit Sub
        End If
        
        'Now you can be atacked
        .flags.NoPuedeSerAtacado = False
        
        If Spell < 1 Then
            .flags.Hechizo = 0
            Exit Sub
        ElseIf Spell > MAXUSERHECHIZOS Then
            .flags.Hechizo = 0
            Exit Sub

        End If
        
        .flags.Hechizo = .Stats.UserHechizos(Spell)

    End With

End Sub

''
' Handles the "LeftClick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleLeftClick(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex).incomingData
        'Remove packet ID
        Call .ReadByte
        
        Dim X As Integer

        Dim Y As Integer
        
        X = .ReadInteger()
        Y = .ReadInteger()
        
        Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, X, Y)

    End With

End Sub

''
' Handles the "AccionClick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleAccionClick(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex).incomingData
        'Remove packet ID
        Call .ReadByte
        
        Dim X As Integer

        Dim Y As Integer
        
        X = .ReadInteger()
        Y = .ReadInteger()
        
        Call Accion(UserIndex, UserList(UserIndex).Pos.Map, X, Y)

    End With

End Sub

''
' Handles the "Work" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWork(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 13/01/2010 (ZaMa)
    '13/01/2010: ZaMa - El pirata se puede ocultar en barca
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Skill As eSkill
        
        Skill = .incomingData.ReadByte()
        
        If UserList(UserIndex).flags.Muerto = 1 Then Exit Sub
        
        'If exiting, cancel
        Call CancelExit(UserIndex)
        
        'Si esta casteando, lo cancelamos
        Call CancelCast(UserIndex)
        
        Select Case Skill
        
            Case Robar, Magia, Domar
                Call WriteMultiMessage(UserIndex, eMessages.WorkRequestTarget, Skill)
                
            Case Ocultarse
                
                ' Verifico si se peude ocultar en este mapa
                If MapZonas(.Pos.Map, UserZonaId(UserIndex)).OcultarSinEfecto = 1 Then
                    Call WriteConsoleMsg(UserIndex, "Ocultarse no funciona aqui!", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
                
                If .flags.EnConsulta Then
                    Call WriteConsoleMsg(UserIndex, "No puedes ocultarte si estas en consulta.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
            
                If .flags.Navegando = 1 Then
                    If .clase <> eClass.Pirat Then

                        '[CDT 17-02-2004]
                        If Not .flags.UltimoMensaje = 3 Then
                            Call WriteConsoleMsg(UserIndex, "No puedes ocultarte si estas navegando.", FontTypeNames.FONTTYPE_INFO)
                            .flags.UltimoMensaje = 3

                        End If

                        '[/CDT]
                        Exit Sub

                    End If

                End If
                
                If .flags.Oculto = 1 Then

                    '[CDT 17-02-2004]
                    If Not .flags.UltimoMensaje = 2 Then
                        Call WriteConsoleMsg(UserIndex, "Ya estas oculto.", FontTypeNames.FONTTYPE_INFO)
                        .flags.UltimoMensaje = 2

                    End If

                    '[/CDT]
                    Exit Sub

                End If
                
                Call DoOcultarse(UserIndex)
                
        End Select
        
    End With

End Sub

''
' Handles the "UseSpellMacro" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUseSpellMacro(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Call SendData(SendTarget.ToAdmins, UserIndex, PrepareMessageConsoleMsg(.Name & " fue expulsado por Anti-macro de hechizos.", FontTypeNames.FONTTYPE_FIGHT))
        Call WriteErrorMsg(UserIndex, "Has sido expulsado por usar macro de hechizos. Recomendamos leer el reglamento sobre el tema macros.")
        Call CloseUser(UserIndex)

    End With

End Sub

''
' Handles the "UseItem" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUseItem(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Slot As Byte
        
        Slot = .incomingData.ReadByte()
        
        If Slot <= .CurrentInventorySlots And Slot > 0 Then
            If .Invent.Object(Slot).ObjIndex = 0 Then Exit Sub

        End If
        
        If .flags.Meditando Then

            Exit Sub    'The error message should have been provided by the client.

        End If
        
        Call UseInvItem(UserIndex, Slot)

    End With

End Sub

''
' Handles the "CraftearItem" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCraftearItem(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Item As Long
        Dim Cantidad As Integer
        Dim Profesion As Byte
        
        Item = .incomingData.ReadLong()
        Cantidad = .incomingData.ReadInteger()
        Profesion = .incomingData.ReadByte()
        
        If Item < 1 Or Cantidad < 1 Then Exit Sub
        
        Call ComenzarCrafteo(UserIndex, Item, Cantidad, Profesion)
    End With

End Sub

Private Sub HandleWorkClose(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Lorwik
    'Last Modification: 21/08/2020
    '
    '***************************************************
    With UserList(UserIndex)
    
        'Remove packet ID
        Call .incomingData.ReadByte

        .flags.Trabajando = 0
    
    End With
    
    
End Sub

''
' Handles the "CraftCarpenter" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCraftCarpenter(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Item As Integer
        Dim Cantidad As Integer
        
        Item = .incomingData.ReadInteger()
        Cantidad = .incomingData.ReadInteger()
        
        If Item < 1 Then Exit Sub
        
        If ObjData(Item).SkCarpinteria = 0 Then Exit Sub
        
        If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub
        'Comprobamos que no se encuentra trabajando, para prevenir bugs y hacks
        If .flags.MacroTrabajo = 0 Then
            .flags.MacroTrabajaObj = Item
            .flags.MacroCountObj = Cantidad
            .flags.MacroTrabajo = eMacroTrabajo.Carpinteando
            Call WriteConsoleMsg(UserIndex, "Comienzas a trabajar.", FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(UserIndex, "Ya te encuentras trabajando.", FontTypeNames.FONTTYPE_INFO)
        End If

    End With
    
errHandler:
    Call LogError("Error en HandleCraftcarpenter en " & Erl & " - Item: " & Item & ". Err " & Err.Number & " " & Err.description)

End Sub

''
' Handles the "WorkLeftClick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWorkLeftClick(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 14/01/2010 (ZaMa)
    '16/11/2009: ZaMa - Agregada la posibilidad de extraer madera elfica.
    '12/01/2010: ZaMa - Ahora se admiten armas arrojadizas (proyectiles sin municiones).
    '14/01/2010: ZaMa - Ya no se pierden municiones al atacar npcs con dueno.
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim X           As Integer

        Dim Y           As Integer

        Dim Skill       As eSkill

        Dim DummyINT    As Integer

        Dim tU          As Integer   'Target user

        Dim tN          As Integer   'Target NPC
        
        Dim WeaponIndex As Integer
        
        X = .incomingData.ReadInteger()
        Y = .incomingData.ReadInteger()
        
        Skill = .incomingData.ReadByte()
        
        If .flags.Muerto = 1 Or .flags.Descansar Or .flags.Meditando Or Not InMapBounds(.Pos.Map, X, Y) Then Exit Sub

        If Not InRangoVision(UserIndex, X, Y) Then
            Call WritePosUpdate(UserIndex)
            Exit Sub

        End If
        
        'If exiting, cancel
        Call CancelExit(UserIndex)
        
        'Si esta casteando, lo cancelamos
        Call CancelCast(UserIndex)
        
        Select Case Skill

            Case eSkill.Proyectiles
                
                'Check attack interval
                If Not IntervaloPermiteAtacar(UserIndex, False) Then Exit Sub

                'Check Magic interval
                If Not IntervaloPermiteLanzarSpell(UserIndex, False) Then Exit Sub

                'Check bow's interval
                If Not IntervaloPermiteUsarArcos(UserIndex) Then Exit Sub
                
                Call LanzarProyectil(UserIndex, X, Y)
                            
            Case eSkill.Magia

                'Check the map allows spells to be casted.
                If MapZonas(.Pos.Map, UserZonaId(UserIndex)).MagiaSinEfecto > 0 Then
                    Call WriteConsoleMsg(UserIndex, "Una fuerza oscura te impide canalizar tu energia.", FontTypeNames.FONTTYPE_FIGHT)
                    Exit Sub
                End If
                
                'Target whatever is in that tile
                Call LookatTile(UserIndex, .Pos.Map, X, Y)
                
                'If it's outside range log it and exit
                If Abs(.Pos.X - X) > RANGO_VISION_X Or Abs(.Pos.Y - Y) > RANGO_VISION_Y Then
                    Call LogCheating("Ataque fuera de rango de " & .Name & "(" & .Pos.Map & "/" & .Pos.X & "/" & .Pos.Y & ") ip: " & .IP & " a la posicion (" & .Pos.Map & "/" & X & "/" & Y & ")")
                    Exit Sub
                End If
                
                'Check bow's interval
                If Not IntervaloPermiteUsarArcos(UserIndex, False) Then Exit Sub
                
                'Check Spell-Hit interval
                If Not IntervaloPermiteGolpeMagia(UserIndex) Then

                    'Check Magic interval
                    If Not IntervaloPermiteLanzarSpell(UserIndex) Then
                        Exit Sub
                    End If

                End If
                
                'Check intervals and cast
                If .flags.Hechizo > 0 Then
                    Call LanzarHechizo(.flags.Hechizo, UserIndex)
                    
                Else
                    Call WriteConsoleMsg(UserIndex, "Primero selecciona el hechizo que quieres lanzar!", FontTypeNames.FONTTYPE_INFO)

                End If
            
            Case eSkill.Robar

                'Does the map allow us to steal here?
                If MapZonas(.Pos.Map, UserZonaId(UserIndex)).Pk Then
                    
                    'Check interval
                    If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub
                    
                    'Target whatever is in that tile
                    Call LookatTile(UserIndex, UserList(UserIndex).Pos.Map, X, Y)
                    
                    tU = .flags.TargetUser
                    
                    If tU > 0 And tU <> UserIndex Then

                        'Can't steal administrative players
                        If UserList(tU).flags.Privilegios And PlayerType.User Then
                            If UserList(tU).flags.Muerto = 0 Then
                                If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 2 Then
                                    Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                                    Exit Sub

                                End If
                                 
                                '17/09/02
                                'Check the trigger
                                If MapData(UserList(tU).Pos.Map, X, Y).Trigger = eTrigger.ZONASEGURA Then
                                    Call WriteConsoleMsg(UserIndex, "No puedes robar aqui.", FontTypeNames.FONTTYPE_WARNING)
                                    Exit Sub

                                End If
                                 
                                If MapData(.Pos.Map, .Pos.X, .Pos.Y).Trigger = eTrigger.ZONASEGURA Then
                                    Call WriteConsoleMsg(UserIndex, "No puedes robar aqui.", FontTypeNames.FONTTYPE_WARNING)
                                    Exit Sub

                                End If
                                 
                                Call DoRobar(UserIndex, tU)

                            End If

                        End If

                    Else
                        Call WriteConsoleMsg(UserIndex, "No hay a quien robarle!", FontTypeNames.FONTTYPE_INFO)

                    End If

                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes robar en zonas seguras!", FontTypeNames.FONTTYPE_INFO)

                End If
                
            Case eSkill.pesca
                WeaponIndex = .Invent.WeaponEqpObjIndex

                If WeaponIndex = 0 Then Exit Sub
                
                'Check interval
                If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub
                
                If MapData(.Pos.Map, .Pos.X, .Pos.Y).Trigger = eTrigger.BAJOTECHO Or MapData(.Pos.Map, .Pos.X, .Pos.Y).Trigger = eTrigger.CASA Then
                    Call WriteConsoleMsg(UserIndex, "No puedes pescar desde donde te encuentras.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
                
                If HayAgua(.Pos.Map, X, Y) Then
                
                    If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 2 Then
                        Call WriteConsoleMsg(UserIndex, "No puedes pescar desde donde te encuentras.", FontTypeNames.FONTTYPE_INFO)
                        Exit Sub
                    End If

                    Select Case WeaponIndex

                        Case CANA_PESCA
                            .flags.MacroTrabajo = eMacroTrabajo.Pescar
                        
                        Case RED_PESCA
                        
                            If .Stats.UserSkills(eSkill.pesca) < ObjData(WeaponIndex).MinSkill Then
                                Call WriteConsoleMsg(UserIndex, "No tienes conocimientos en Pesca suficiente para usar la red. Necesitas al menos " & ObjData(WeaponIndex).MinSkill & " Skills.", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
                            End If
                            
                            If .flags.Navegando = 0 Then
                                Call WriteConsoleMsg(UserIndex, "Para pescar necesitas estar en una barca.", FontTypeNames.FONTTYPE_INFO)
                                Exit Sub
                            End If
                                              
                            .flags.MacroTrabajo = eMacroTrabajo.PescarRed
   
                        Case Else

                            Exit Sub    'Invalid item!

                    End Select
                    
                    Call WriteConsoleMsg(UserIndex, "Comienzas a trabajar.", FontTypeNames.FONTTYPE_INFO)
                    
                Else
                    Call WriteConsoleMsg(UserIndex, "No hay agua donde pescar. Busca un lago, rio o mar.", FontTypeNames.FONTTYPE_INFO)

                End If
            
            Case eSkill.Domar
                'Modificado 25/11/02
                'Optimizado y solucionado el bug de la doma de
                'criaturas hostiles.
                
                'Target whatever is that tile
                Call LookatTile(UserIndex, .Pos.Map, X, Y)
                tN = .flags.TargetNPC
                
                If tN > 0 Then
                    If Npclist(tN).flags.Domable > 0 Then
                        If Abs(.Pos.X - X) + Abs(.Pos.Y - Y) > 2 Then
                            Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
                            Exit Sub

                        End If
                        
                        If LenB(Npclist(tN).flags.AttackedBy) <> 0 Then
                            Call WriteConsoleMsg(UserIndex, "No puedes domar una criatura que esta luchando con un jugador.", FontTypeNames.FONTTYPE_INFO)
                            Exit Sub

                        End If
                        
                        Call DoDomar(UserIndex, tN)
                    Else
                        Call WriteConsoleMsg(UserIndex, "No puedes domar a esa criatura.", FontTypeNames.FONTTYPE_INFO)

                    End If

                Else
                    Call WriteConsoleMsg(UserIndex, "No hay ninguna criatura alli!", FontTypeNames.FONTTYPE_INFO)

                End If
            
            Case FundirMetal    'UGLY!!! This is a constant, not a skill!!
                If PuedeLingotear(UserIndex) Then
                    .flags.MacroTrabajo = eMacroTrabajo.Lingotear
                    Call WriteConsoleMsg(UserIndex, "Comienzas a trabajar.", FontTypeNames.FONTTYPE_INFO)
                End If
            
            Case eSkill.herreria
                'Target wehatever is in that tile
                Call LookatTile(UserIndex, .Pos.Map, X, Y)
                
                If ConoceProfesion(UserIndex, eSkill.herreria) < 0 Then
                    Call WriteConsoleMsg(UserIndex, "No conoces esa profesion.", FontTypeNames.FONTTYPE_INFOBOLD)
                    Exit Sub
                End If
                
                If .flags.TargetObj > 0 Then
                    If ObjData(.flags.TargetObj).OBJType = eOBJType.otYunque Then
                        Call WriteInitTrabajo(UserIndex, eSkill.herreria)
                        
                    Else
                        Call WriteConsoleMsg(UserIndex, "Ahi no hay ningUn yunque.", FontTypeNames.FONTTYPE_INFO)

                    End If

                Else
                    Call WriteConsoleMsg(UserIndex, "Ahi no hay ningUn yunque.", FontTypeNames.FONTTYPE_INFO)

                End If

        End Select

    End With

End Sub

''
' Handles the "InvitarPartyClick" message.

Private Sub HandleInvitarPartyClick(ByVal UserIndex As Integer)
'***************************************************
'Author: Lorwik
'Last Modification: 05/11/2020
'***************************************************
    
    With UserList(UserIndex)
    
        If .incomingData.Length < 3 Then
            Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
            Exit Sub
    
        End If
        
        Dim X           As Integer
    
        Dim Y           As Integer
    
        Dim Aleatorio   As Integer
        
        'Remove packet ID
        Call .incomingData.ReadByte
            
        X = .incomingData.ReadInteger()
        Y = .incomingData.ReadInteger()
    
        If .flags.Muerto = 1 Or .flags.Descansar Or .flags.Meditando Or Not InMapBounds(.Pos.Map, X, Y) Then Exit Sub
    
        If Not InRangoVision(UserIndex, X, Y) Then
            Call WritePosUpdate(UserIndex)
            Exit Sub
    
        End If
    
        'If exiting, cancel
        Call CancelExit(UserIndex)
            
        'Si esta casteando, lo cancelamos
        Call CancelCast(UserIndex)
    
        'Target whatever is in that tile
        Call LookatTile(UserIndex, .Pos.Map, X, Y)
                    
        If .flags.TargetUser <= 0 Then Exit Sub
                    
        'If it's outside range log it and exit
        If Abs(.Pos.X - X) > RANGO_VISION_X Or Abs(.Pos.Y - Y) > RANGO_VISION_Y Then
            Call LogCheating("Ataque fuera de rango de " & .Name & "(" & .Pos.Map & "/" & .Pos.X & "/" & .Pos.Y & ") ip: " & .IP & " a la posicion (" & .Pos.Map & "/" & X & "/" & Y & ")")
            Exit Sub
        End If
        
        '¿Se invita a si mismo?
        If UserIndex = .flags.TargetUser Then Exit Sub
        
        '¿No tengo grupo?
        If .PartyIndex = 0 Then
            
            '¿El otro tampoco tiene?
            If UserList(.flags.TargetUser).PartyIndex = 0 Then
            
                'Lo podemos crear?
                If Not mdParty.PuedeCrearParty(UserIndex) Then Exit Sub
                
                '¿Estan creando grupo?
                If .FormandoGrupo <> .ID Then
                    
                    .FormandoGrupo = UserList(.flags.TargetUser).ID
                    UserList(.flags.TargetUser).FormandoGrupo = UserList(.flags.TargetUser).ID 'Se anota asi mismo, señal que es el invitado
                    
                    Call WriteConsoleMsg(UserIndex, "Has enviado una peticion a " & UserList(.flags.TargetUser).Name & " para crear un grupo.", FontTypeNames.FONTTYPE_INFO)
                    Call WriteConsoleMsg(.flags.TargetUser, UserList(UserIndex).Name & " te ha invitado para crear un grupo.", FontTypeNames.FONTTYPE_INFO)
                    
                Else '¿Es la respuesta?
                    'Lo creamos
                    Call mdParty.CrearParty(UserIndex)
                    
                    'Metemos al target
                    UserList(.flags.TargetUser).PartySolicitud = .PartyIndex
                    
                    'Lo aceptamos
                    Call mdParty.AprobarIngresoAParty(UserIndex, .flags.TargetUser)
                    
                    .FormandoGrupo = 0
                    UserList(.flags.TargetUser).FormandoGrupo = 0
                    
                    Exit Sub
                End If
                
            Else '¿El otro SI tiene grupo?
            
                '¿Es el lider?
                If Parties(UserList(.flags.TargetUser).PartyIndex).EsPartyLeader(.flags.TargetUser) Then
                    'Enviamos peticion para unirme
                    Call mdParty.SolicitarIngresoAParty(UserIndex)
                    Exit Sub
                Else '¿No lo es?
                    Call WriteConsoleMsg(UserIndex, UserList(.flags.TargetUser).Name & " ya pertenece a un grupo.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
            
            End If
            
        Else '¿SI tengo party?

            '¿Soy el lider?
            If Parties(.PartyIndex).EsPartyLeader(UserIndex) Then
                '¿Solicito entrar a mi party?
                If UserList(.flags.TargetUser).PartySolicitud = .PartyIndex Then
                    'Lo aceptamos
                    Call mdParty.AprobarIngresoAParty(UserIndex, .flags.TargetUser)
                    
                Else '¿no?
                    Call WriteConsoleMsg(UserIndex, UserList(.flags.TargetUser).Name & " no ha solicitado entrar a tu grupo.", FontTypeNames.FONTTYPE_PARTY)
                    
                End If
            End If
        
        End If
    
    End With
    
End Sub

''
' Handles the "CreateNewGuild" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCreateNewGuild(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/11/09
    '05/11/09: Pato - Ahora se quitan los espacios del principio y del fin del nombre del clan
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 9 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim Desc      As String

        Dim GuildName As String

        Dim Site      As String

        Dim codex()   As String

        Dim errorStr  As String
        
        Desc = Buffer.ReadASCIIString()
        GuildName = Trim$(Buffer.ReadASCIIString())
        Site = Buffer.ReadASCIIString()
        codex = Split(Buffer.ReadASCIIString(), SEPARATOR)
        
        If modGuilds.CrearNuevoClan(UserIndex, Desc, GuildName, Site, codex, .FundandoGuildAlineacion, errorStr) Then
            Dim Message As String
            Message = .Name & " fundo el clan " & GuildName & " de alineacion " & modGuilds.GuildAlignment(.GuildIndex)

            Call SendData(SendTarget.Toall, UserIndex, PrepareMessageConsoleMsg(Message, FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.Toall, 0, PrepareMessagePlayWave(44, NO_3D_SOUND, NO_3D_SOUND))
            
            'Update tag
            Call RefreshCharStatus(UserIndex)

            'Aqui solo vamos a hacer un request a los endpoints de la aplicacion en Node.js
            'el repositorio para hacer funcionar esto, es este: https://github.com/ao-libre/ao-api-server
            'Si no tienen interes en usarlo pueden desactivarlo en el Server.ini
            If ConexionAPI Then
                Call ApiEndpointSendNewGuildCreatedMessageDiscord(Message, Desc, GuildName, Site)
            End If
        Else
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "SpellInfo" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSpellInfo(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    If UserList(UserIndex).incomingData.Length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim spellSlot As Byte
        Dim Spell As Integer
        
        spellSlot = .incomingData.ReadByte()
        
        'Validate slot
        If spellSlot < 1 Or spellSlot > MAXUSERHECHIZOS Then
            Call WriteConsoleMsg(UserIndex, "¡Primero selecciona el hechizo.!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Validate spell in the slot
        Spell = .Stats.UserHechizos(spellSlot)
        If Spell > 0 And Spell < NumeroHechizos + 1 Then
            With Hechizos(Spell)
                'Send information
                Call WriteConsoleMsg(UserIndex, "%%%%%%%%%%%% INFO DEL HECHIZO %%%%%%%%%%%%" & vbCrLf _
                                               & "Nombre:" & .Nombre & vbCrLf _
                                               & "Descripción:" & .Desc & vbCrLf _
                                               & "Skill requerido: " & .MinSkill & " de magia." & vbCrLf _
                                               & "Mana necesario: " & .ManaRequerido & vbCrLf _
                                               & "Stamina necesaria: " & .StaRequerido & vbCrLf _
                                               & "%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%", FontTypeNames.FONTTYPE_INFO)
            End With
        End If
    End With
End Sub

''
' Handles the "EquipItem" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleEquipItem(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim itemSlot As Byte
        
        itemSlot = .incomingData.ReadByte()
        
        'Dead users can't equip items
        If .flags.Muerto = 1 Then Exit Sub
        
        'Validate item slot
        If itemSlot > .CurrentInventorySlots Or itemSlot < 1 Then Exit Sub
        
        If .Invent.Object(itemSlot).ObjIndex = 0 Then Exit Sub
        
        Call EquiparInvItem(UserIndex, itemSlot)

    End With

End Sub

''
' Handles the "ChangeHeading" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleChangeHeading(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 06/28/2008
    'Last Modified By: NicoNZ
    ' 10/01/2008: Tavo - Se cancela la salida del juego si el user esta saliendo
    ' 06/28/2008: NicoNZ - Solo se puede cambiar si esta inmovilizado.
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Heading As eHeading

        Dim posX    As Integer

        Dim posY    As Integer
                
        Heading = .incomingData.ReadByte()
        
        If .flags.Paralizado = 1 And .flags.Inmovilizado = 0 Then

            Select Case Heading

                Case eHeading.NORTH
                    posY = -1

                Case eHeading.EAST
                    posX = 1

                Case eHeading.SOUTH
                    posY = 1

                Case eHeading.WEST
                    posX = -1

            End Select
            
            If LegalPos(.Pos.Map, .Pos.X + posX, .Pos.Y + posY, CBool(.flags.Navegando), Not CBool(.flags.Navegando)) Then
                Exit Sub

            End If

        End If
        
        'Validate heading (VB won't say invalid cast if not a valid index like .Net languages would do... *sigh*)
        If Heading > 0 And Heading < 5 Then
            .Char.Heading = Heading
            Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraAnim, .Char.AuraColor)

        End If

    End With

End Sub

''
' Handles the "Train" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTrain(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim SpawnedNpc As Integer

        Dim PetIndex   As Byte
        
        PetIndex = .incomingData.ReadByte()
        
        If .flags.TargetNPC = 0 Then Exit Sub
        
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Entrenador Then Exit Sub
        
        If Npclist(.flags.TargetNPC).Mascotas < MAXMASCOTASENTRENADOR Then
            If PetIndex > 0 And PetIndex < Npclist(.flags.TargetNPC).NroCriaturas + 1 Then
                'Create the creature
                SpawnedNpc = SpawnNpc(Npclist(.flags.TargetNPC).Criaturas(PetIndex).NPCIndex, Npclist(.flags.TargetNPC).Pos, True, False)
                
                If SpawnedNpc > 0 Then
                    Npclist(SpawnedNpc).MaestroNpc = .flags.TargetNPC
                    Npclist(.flags.TargetNPC).Mascotas = Npclist(.flags.TargetNPC).Mascotas + 1

                End If

            End If

        Else
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("No puedo traer mas criaturas, mata las existentes.", Npclist(.flags.TargetNPC).Char.CharIndex, 255, 255, 255))

        End If

    End With

End Sub

''
' Handles the "CommerceBuy" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCommerceBuy(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Slot   As Byte

        Dim Amount As Integer
        
        Slot = .incomingData.ReadByte()
        Amount = .incomingData.ReadInteger()
        
        'Dead people can't commerce...
        If .flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "Estas muerto!!", FontTypeNames.FONTTYPE_INFO)
            Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
            Exit Sub

        End If
        
        'El target es un NPC valido?
        If .flags.TargetNPC < 1 Then Exit Sub
            
        'El NPC puede comerciar?
        If Npclist(.flags.TargetNPC).Comercia = 0 Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("No tengo ningun interes en comerciar.", Npclist(.flags.TargetNPC).Char.CharIndex, 255, 255, 255))
            Exit Sub

        End If
        
        'Only if in commerce mode....
        If Not .flags.Comerciando Then
            Call WriteConsoleMsg(UserIndex, "No estas comerciando.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        'User compra el item
        Call Comercio(eModoComercio.Compra, UserIndex, .flags.TargetNPC, Slot, Amount)

    End With

End Sub

''
' Handles the "BankExtractItem" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankExtractItem(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Slot   As Byte

        Dim Amount As Integer
        
        Slot = .incomingData.ReadByte()
        Amount = .incomingData.ReadInteger()
        
        'Dead people can't commerce
        If .flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "Estas muerto!!", FontTypeNames.FONTTYPE_INFO)
            Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
            Exit Sub

        End If
        
        'El target es un NPC valido?
        If .flags.TargetNPC < 1 Then Exit Sub
        
        'Es el banquero?
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Banquero Then
            Exit Sub

        End If
        
        'User retira el item del slot
        Call UserRetiraItem(UserIndex, Slot, Amount)

    End With

End Sub

''
' Handles the "CommerceSell" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCommerceSell(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Slot   As Byte

        Dim Amount As Integer
        
        Slot = .incomingData.ReadByte()
        Amount = .incomingData.ReadInteger()
        
        'Dead people can't commerce...
        If .flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "Estas muerto!!", FontTypeNames.FONTTYPE_INFO)
            Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
            Exit Sub

        End If
        
        'El target es un NPC valido?
        If .flags.TargetNPC < 1 Then Exit Sub
        
        'El NPC puede comerciar?
        If Npclist(.flags.TargetNPC).Comercia = 0 Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("No tengo ningun interes en comerciar.", Npclist(.flags.TargetNPC).Char.CharIndex, 255, 255, 255))
            Exit Sub

        End If
        
        'User compra el item del slot
        Call Comercio(eModoComercio.Venta, UserIndex, .flags.TargetNPC, Slot, Amount)

    End With

End Sub

''
' Handles the "BankDeposit" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankDeposit(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Slot   As Byte

        Dim Amount As Integer
        
        Slot = .incomingData.ReadByte()
        Amount = .incomingData.ReadInteger()
        
        'Dead people can't commerce...
        If .flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "Estas muerto!!", FontTypeNames.FONTTYPE_INFO)
            Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
            Exit Sub

        End If
        
        'El target es un NPC valido?
        If .flags.TargetNPC < 1 Then Exit Sub
        
        'El NPC puede comerciar?
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Banquero Then
            Exit Sub

        End If
        
        'User deposita el item del slot rdata
        Call UserDepositaItem(UserIndex, Slot, Amount)

    End With

End Sub

''
' Handles the "ForumPost" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleForumPost(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 02/01/2010
    '02/01/2010: ZaMa - Implemento nuevo sistema de foros
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 6 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim ForumMsgType As eForumMsgType
        
        Dim file         As String

        Dim Title        As String

        Dim Post         As String

        Dim ForumIndex   As Integer

        Dim postFile     As String

        Dim ForumType    As Byte
                
        ForumMsgType = Buffer.ReadByte()
        
        Title = Buffer.ReadASCIIString()
        Post = Buffer.ReadASCIIString()
        
        If .flags.TargetObj > 0 Then
            ForumType = ForumAlignment(ForumMsgType)
            
            Select Case ForumType
            
                Case eForumType.ieGeneral
                    ForumIndex = GetForumIndex(ObjData(.flags.TargetObj).ForoID)
                    
                Case eForumType.ieREAL
                    ForumIndex = GetForumIndex(FORO_REAL_ID)
                    
                Case eForumType.ieCAOS
                    ForumIndex = GetForumIndex(FORO_CAOS_ID)
                    
            End Select
            
            Call AddPost(ForumIndex, Post, .Name, Title, EsAnuncio(ForumMsgType))

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "MoveSpell" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleMoveSpell(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex).incomingData
        'Remove packet ID
        Call .ReadByte
        
        Dim Dir As Integer
        
        If .ReadBoolean() Then
            Dir = 1
        Else
            Dir = -1

        End If
        
        Call DesplazarHechizo(UserIndex, Dir, .ReadByte())

    End With

End Sub

''
' Handles the "MoveBank" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleMoveBank(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Torres Patricio (Pato)
    'Last Modification: 06/14/09
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex).incomingData
        'Remove packet ID
        Call .ReadByte
        
        Dim Dir      As Integer

        Dim Slot     As Byte

        Dim TempItem As obj
        
        If .ReadBoolean() Then
            Dir = 1
        Else
            Dir = -1

        End If
        
        Slot = .ReadByte()

    End With
        
    With UserList(UserIndex)
        TempItem.ObjIndex = .BancoInvent.Object(Slot).ObjIndex
        TempItem.Amount = .BancoInvent.Object(Slot).Amount
        
        If Dir = 1 Then 'Mover arriba
            .BancoInvent.Object(Slot) = .BancoInvent.Object(Slot - 1)
            .BancoInvent.Object(Slot - 1).ObjIndex = TempItem.ObjIndex
            .BancoInvent.Object(Slot - 1).Amount = TempItem.Amount

            Call UpdateBanUserInv(False, UserIndex, Slot - 1)
        Else 'mover abajo
            .BancoInvent.Object(Slot) = .BancoInvent.Object(Slot + 1)
            .BancoInvent.Object(Slot + 1).ObjIndex = TempItem.ObjIndex
            .BancoInvent.Object(Slot + 1).Amount = TempItem.Amount

            Call UpdateBanUserInv(False, UserIndex, Slot + 1)
        End If

        Call UpdateBanUserInv(False, UserIndex, Slot)

    End With

End Sub

''
' Handles the "ClanCodexUpdate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleClanCodexUpdate(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim Desc    As String

        Dim codex() As String
        
        Desc = Buffer.ReadASCIIString()
        codex = Split(Buffer.ReadASCIIString(), SEPARATOR)
        
        Call modGuilds.ChangeCodexAndDesc(Desc, codex, .GuildIndex)
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "UserCommerceOffer" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUserCommerceOffer(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 24/11/2009
    '24/11/2009: ZaMa - Nuevo sistema de comercio
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 7 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Amount    As Long

        Dim Slot      As Byte

        Dim tUser     As Integer

        Dim OfferSlot As Byte

        Dim ObjIndex  As Integer
        
        Slot = .incomingData.ReadByte()
        Amount = .incomingData.ReadLong()
        OfferSlot = .incomingData.ReadByte()
        
        'Get the other player
        tUser = .ComUsu.DestUsu
        
        ' If he's already confirmed his offer, but now tries to change it, then he's cheating
        If UserList(UserIndex).ComUsu.Confirmo = True Then
            
            ' Finish the trade
            Call FinComerciarUsu(UserIndex)
        
            If tUser <= 0 Or tUser > MaxUsers Then
                Call FinComerciarUsu(tUser)

            End If
        
            Exit Sub

        End If
        
        'If slot is invalid and it's not gold or it's not 0 (Substracting), then ignore it.
        If ((Slot < 0 Or Slot > UserList(UserIndex).CurrentInventorySlots) And Slot <> FLAGORO) Then Exit Sub
        
        'If OfferSlot is invalid, then ignore it.
        If OfferSlot < 1 Or OfferSlot > MAX_OFFER_SLOTS + 1 Then Exit Sub
        
        ' Can be negative if substracted from the offer, but never 0.
        If Amount = 0 Then Exit Sub
        
        'Has he got enough??
        If Slot = FLAGORO Then

            ' Can't offer more than he has
            If Amount > .Stats.Gld - .ComUsu.GoldAmount Then
                Call WriteCommerceChat(UserIndex, "No tienes esa cantidad de oro para agregar a la oferta.", FontTypeNames.FONTTYPE_TALK)
                Exit Sub

            End If
            
            If Amount < 0 Then
                If Abs(Amount) > .ComUsu.GoldAmount Then
                    Amount = .ComUsu.GoldAmount * (-1)

                End If

            End If

        Else

            'If modifing a filled offerSlot, we already got the objIndex, then we don't need to know it
            If Slot <> 0 Then ObjIndex = .Invent.Object(Slot).ObjIndex

            ' Can't offer more than he has
            If Not HasEnoughItems(UserIndex, ObjIndex, TotalOfferItems(ObjIndex, UserIndex) + Amount) Then
                
                Call WriteCommerceChat(UserIndex, "No tienes esa cantidad.", FontTypeNames.FONTTYPE_TALK)
                Exit Sub

            End If
            
            If Amount < 0 Then
                If Abs(Amount) > .ComUsu.cant(OfferSlot) Then
                    Amount = .ComUsu.cant(OfferSlot) * (-1)

                End If

            End If
        
            'No se puede comerciar con los items de newbie
            If ItemNewbie(ObjIndex) Then
                Call WriteCancelOfferItem(UserIndex, OfferSlot)
                Exit Sub

            End If
            
            'No se puede comerciar con la runa de hogar
            If ObjData(ObjIndex).OBJType = otRunaHogar Then
                Call WriteCancelOfferItem(UserIndex, OfferSlot)
                Exit Sub

            End If
            
            'Don't allow to sell boats if they are equipped (you can't take them off in the water and causes trouble)
            If .flags.Navegando = 1 Then
                If .Invent.BarcoSlot = Slot Then
                    Call WriteCommerceChat(UserIndex, "No puedes vender tu barco mientras lo estes usando.", FontTypeNames.FONTTYPE_TALK)
                    Exit Sub

                End If

            End If
            
            If .flags.Equitando = 1 Then
                If .Invent.MonturaEqpSlot = Slot Then
                    Call WriteConsoleMsg(UserIndex, "No podes vender tu montura mientras lo estes usando.", FontTypeNames.FONTTYPE_TALK)
                    Exit Sub
                End If
            End If

            If .Invent.MochilaEqpSlot > 0 Then
                If .Invent.MochilaEqpSlot = Slot Then
                    Call WriteCommerceChat(UserIndex, "No puedes vender tu alforja o mochila mientras la estes usando.", FontTypeNames.FONTTYPE_TALK)
                    Exit Sub

                End If

            End If

        End If
        
        Call AgregarOferta(UserIndex, OfferSlot, ObjIndex, Amount, Slot = FLAGORO)
        Call EnviarOferta(tUser, OfferSlot)

    End With

End Sub

''
' Handles the "GuildAcceptPeace" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildAcceptPeace(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim Guild          As String

        Dim errorStr       As String

        Dim otherClanIndex As String
        
        Guild = Buffer.ReadASCIIString()
        
        otherClanIndex = modGuilds.r_AceptarPropuestaDePaz(UserIndex, Guild, errorStr)
        
        If otherClanIndex = 0 Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Tu clan ha firmado la paz con " & Guild & ".", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageConsoleMsg("Tu clan ha firmado la paz con " & modGuilds.GuildName(.GuildIndex) & ".", FontTypeNames.FONTTYPE_GUILD))

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "GuildRejectAlliance" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildRejectAlliance(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim Guild          As String

        Dim errorStr       As String

        Dim otherClanIndex As String
        
        Guild = Buffer.ReadASCIIString()
        
        otherClanIndex = modGuilds.r_RechazarPropuestaDeAlianza(UserIndex, Guild, errorStr)
        
        If otherClanIndex = 0 Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Tu clan rechazado la propuesta de alianza de " & Guild, FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageConsoleMsg(modGuilds.GuildName(.GuildIndex) & " ha rechazado nuestra propuesta de alianza con su clan.", FontTypeNames.FONTTYPE_GUILD))

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "GuildRejectPeace" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildRejectPeace(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim Guild          As String

        Dim errorStr       As String

        Dim otherClanIndex As String
        
        Guild = Buffer.ReadASCIIString()
        
        otherClanIndex = modGuilds.r_RechazarPropuestaDePaz(UserIndex, Guild, errorStr)
        
        If otherClanIndex = 0 Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Tu clan rechazado la propuesta de paz de " & Guild & ".", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageConsoleMsg(modGuilds.GuildName(.GuildIndex) & " ha rechazado nuestra propuesta de paz con su clan.", FontTypeNames.FONTTYPE_GUILD))

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "GuildAcceptAlliance" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildAcceptAlliance(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim Guild          As String

        Dim errorStr       As String

        Dim otherClanIndex As String
        
        Guild = Buffer.ReadASCIIString()
        
        otherClanIndex = modGuilds.r_AceptarPropuestaDeAlianza(UserIndex, Guild, errorStr)
        
        If otherClanIndex = 0 Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Tu clan ha firmado la alianza con " & Guild & ".", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherClanIndex, PrepareMessageConsoleMsg("Tu clan ha firmado la paz con " & modGuilds.GuildName(.GuildIndex) & ".", FontTypeNames.FONTTYPE_GUILD))

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "GuildOfferPeace" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildOfferPeace(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim Guild    As String

        Dim proposal As String

        Dim errorStr As String
        
        Guild = Buffer.ReadASCIIString()
        proposal = Buffer.ReadASCIIString()
        
        If modGuilds.r_ClanGeneraPropuesta(UserIndex, Guild, RELACIONES_GUILD.PAZ, proposal, errorStr) Then
            Call WriteConsoleMsg(UserIndex, "Propuesta de paz enviada.", FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "GuildOfferAlliance" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildOfferAlliance(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim Guild    As String

        Dim proposal As String

        Dim errorStr As String
        
        Guild = Buffer.ReadASCIIString()
        proposal = Buffer.ReadASCIIString()
        
        If modGuilds.r_ClanGeneraPropuesta(UserIndex, Guild, RELACIONES_GUILD.ALIADOS, proposal, errorStr) Then
            Call WriteConsoleMsg(UserIndex, "Propuesta de alianza enviada.", FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "GuildAllianceDetails" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildAllianceDetails(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim Guild    As String

        Dim errorStr As String

        Dim details  As String
        
        Guild = Buffer.ReadASCIIString()
        
        details = modGuilds.r_VerPropuesta(UserIndex, Guild, RELACIONES_GUILD.ALIADOS, errorStr)
        
        If LenB(details) = 0 Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteOfferDetails(UserIndex, details)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "GuildPeaceDetails" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildPeaceDetails(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim Guild    As String

        Dim errorStr As String

        Dim details  As String
        
        Guild = Buffer.ReadASCIIString()
        
        details = modGuilds.r_VerPropuesta(UserIndex, Guild, RELACIONES_GUILD.PAZ, errorStr)
        
        If LenB(details) = 0 Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteOfferDetails(UserIndex, details)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "GuildRequestJoinerInfo" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildRequestJoinerInfo(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim User    As String

        Dim details As String
        
        User = Buffer.ReadASCIIString()
        
        details = modGuilds.a_DetallesAspirante(UserIndex, User)
        
        If LenB(details) = 0 Then
            Call WriteConsoleMsg(UserIndex, "El personaje no ha mandado solicitud, o no estas habilitado para verla.", FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteShowUserRequest(UserIndex, details)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "GuildAlliancePropList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildAlliancePropList(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    Call WriteAlianceProposalsList(UserIndex, r_ListaDePropuestas(UserIndex, RELACIONES_GUILD.ALIADOS))

End Sub

''
' Handles the "GuildPeacePropList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildPeacePropList(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    Call WritePeaceProposalsList(UserIndex, r_ListaDePropuestas(UserIndex, RELACIONES_GUILD.PAZ))

End Sub

''
' Handles the "GuildDeclareWar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildDeclareWar(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim Guild           As String

        Dim errorStr        As String

        Dim otherGuildIndex As Integer
        
        Guild = Buffer.ReadASCIIString()
        
        otherGuildIndex = modGuilds.r_DeclararGuerra(UserIndex, Guild, errorStr)
        
        If otherGuildIndex = 0 Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            'WAR shall be!
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("TU CLAN HA ENTRADO EN GUERRA CON " & Guild & ".", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, otherGuildIndex, PrepareMessageConsoleMsg(modGuilds.GuildName(.GuildIndex) & " LE DECLARA LA GUERRA A TU CLAN.", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessagePlayWave(45, NO_3D_SOUND, NO_3D_SOUND))
            Call SendData(SendTarget.ToGuildMembers, otherGuildIndex, PrepareMessagePlayWave(45, NO_3D_SOUND, NO_3D_SOUND))

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "GuildNewWebsite" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildNewWebsite(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Call modGuilds.ActualizarWebSite(UserIndex, Buffer.ReadASCIIString())
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "GuildAcceptNewMember" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildAcceptNewMember(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim errorStr As String

        Dim username As String

        Dim tUser    As Integer
        
        username = Buffer.ReadASCIIString()
        
        If Not modGuilds.a_AceptarAspirante(UserIndex, username, errorStr) Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            tUser = NameIndex(username)

            If tUser > 0 Then
                Call modGuilds.m_ConectarMiembroAClan(tUser, .GuildIndex)
                Call RefreshCharStatus(tUser)

            End If
            
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg(username & " ha sido aceptado como miembro del clan.", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessagePlayWave(43, NO_3D_SOUND, NO_3D_SOUND))

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "GuildRejectNewMember" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildRejectNewMember(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 01/08/07
    'Last Modification by: (liquid)
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim errorStr As String

        Dim username As String

        Dim Reason   As String

        Dim tUser    As Integer
        
        username = Buffer.ReadASCIIString()
        Reason = Buffer.ReadASCIIString()
        
        If Not modGuilds.a_RechazarAspirante(UserIndex, username, errorStr) Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            tUser = NameIndex(username)
            
            If tUser > 0 Then
                Call WriteConsoleMsg(tUser, errorStr & " : " & Reason, FontTypeNames.FONTTYPE_GUILD)
            Else
                'hay que grabar en el char su rechazo
                Call modGuilds.a_RechazarAspiranteChar(username, .GuildIndex, Reason)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "GuildKickMember" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildKickMember(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim username   As String

        Dim GuildIndex As Integer
        
        username = Buffer.ReadASCIIString()
        
        GuildIndex = modGuilds.m_EcharMiembroDeClan(UserIndex, username)
        
        If GuildIndex > 0 Then
            Call SendData(SendTarget.ToGuildMembers, GuildIndex, PrepareMessageConsoleMsg(username & " fue expulsado del clan.", FontTypeNames.FONTTYPE_GUILD))
            Call SendData(SendTarget.ToGuildMembers, GuildIndex, PrepareMessagePlayWave(45, NO_3D_SOUND, NO_3D_SOUND))
        Else
            Call WriteConsoleMsg(UserIndex, "No puedes expulsar ese personaje del clan.", FontTypeNames.FONTTYPE_GUILD)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "GuildUpdateNews" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildUpdateNews(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Call modGuilds.ActualizarNoticias(UserIndex, Buffer.ReadASCIIString())
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "GuildMemberInfo" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildMemberInfo(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Call modGuilds.SendDetallesPersonaje(UserIndex, Buffer.ReadASCIIString())
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "GuildOpenElections" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildOpenElections(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Error As String
        
        If Not modGuilds.v_AbrirElecciones(UserIndex, Error) Then
            Call WriteConsoleMsg(UserIndex, Error, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call SendData(SendTarget.ToGuildMembers, .GuildIndex, PrepareMessageConsoleMsg("Han comenzado las elecciones del clan! Puedes votar escribiendo /VOTO seguido del nombre del personaje, por ejemplo: /VOTO " & .Name, FontTypeNames.FONTTYPE_GUILD))

        End If

    End With

End Sub

''
' Handles the "GuildRequestMembership" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildRequestMembership(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim Guild       As String

        Dim application As String

        Dim errorStr    As String
        
        Guild = Buffer.ReadASCIIString()
        application = Buffer.ReadASCIIString()
        
        If Not modGuilds.a_NuevoAspirante(UserIndex, Guild, application, errorStr) Then
            Call WriteConsoleMsg(UserIndex, errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteConsoleMsg(UserIndex, "Tu solicitud ha sido enviada. Espera prontas noticias del lider de " & Guild & ".", FontTypeNames.FONTTYPE_GUILD)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "GuildRequestDetails" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildRequestDetails(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Call modGuilds.SendGuildDetails(UserIndex, Buffer.ReadASCIIString())
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "Online" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleOnline(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 14/07/19 (Recox)
    'Ahora se muestra una lista de nombres de jugadores online, se suman los gms tambien a la lista (Recox)
    '***************************************************
    Dim i     As Long

    Dim Count As Long
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim UsersNamesOnlines As String

        For i = 1 To LastUser

            If LenB(UserList(i).Name) <> 0 Then

                If i = LastUser Then
                    UsersNamesOnlines = UsersNamesOnlines + UserList(i).Name
                Else
                    UsersNamesOnlines = UsersNamesOnlines + UserList(i).Name + ", "
                End If
                
                Count = Count + 1
            End If

        Next i
        
        Call WriteConsoleMsg(UserIndex, UsersNamesOnlines, FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(UserIndex, "Numero de usuarios: " & CStr(Count), FontTypeNames.FONTTYPE_INFOBOLD)

    End With

    Call WriteConsoleServerUpTimeMsg(UserIndex)

End Sub

''
' Handles the "Quit" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleQuit(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 04/15/2008 (NicoNZ)
    'If user is invisible, it automatically becomes
    'visible before doing the countdown to exit
    '15/04/2008 - No se reseteaban lso contadores de invi ni de ocultar. (NicoNZ)
    '13/01/2020 - Se pusieron nuevas validaciones para las monturas. (Recox)
    '***************************************************
    Dim tUser        As Integer

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        If .flags.Paralizado = 1 Then
            Call WriteConsoleMsg(UserIndex, "No puedes salir estando paralizado.", FontTypeNames.FONTTYPE_WARNING)
            Exit Sub

        End If
        
        'Subastas
        If UserIndex = Subasta.UserIndex Or UserIndex = Subasta.OfertaIndex Then
            Call WriteConsoleMsg(UserIndex, "No puedes salir mientras ofertas en una subasta o realizas una subasta.", FontTypeNames.FONTTYPE_WARNING)
            Exit Sub
        End If
        
        'exit secure commerce
        If .ComUsu.DestUsu > 0 Then
            tUser = .ComUsu.DestUsu
            
            If UserList(tUser).flags.UserLogged Then
                If UserList(tUser).ComUsu.DestUsu = UserIndex Then
                    Call WriteConsoleMsg(tUser, "Comercio cancelado por el otro usuario.", FontTypeNames.FONTTYPE_WARNING)
                    Call FinComerciarUsu(tUser)

                End If

            End If
            
            Call WriteConsoleMsg(UserIndex, "Comercio cancelado.", FontTypeNames.FONTTYPE_WARNING)
            Call FinComerciarUsu(UserIndex)

        End If

        Call Cerrar_Usuario(UserIndex)

    End With

End Sub

''
' Handles the "GuildLeave" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildLeave(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    Dim GuildIndex As Integer
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'obtengo el guildindex
        GuildIndex = m_EcharMiembroDeClan(UserIndex, .Name)
        
        If GuildIndex > 0 Then
            Call WriteConsoleMsg(UserIndex, "Dejas el clan.", FontTypeNames.FONTTYPE_GUILD)
            Call SendData(SendTarget.ToGuildMembers, GuildIndex, PrepareMessageConsoleMsg(.Name & " deja el clan.", FontTypeNames.FONTTYPE_GUILD))
        Else
            Call WriteConsoleMsg(UserIndex, "Tu no puedes salir de este clan.", FontTypeNames.FONTTYPE_GUILD)

        End If

    End With

End Sub

''
' Handles the "RequestAccountState" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestAccountState(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    Dim earnings   As Integer

    Dim Percentage As Integer
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead people can't check their accounts
        If .flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "Estas muerto!!", FontTypeNames.FONTTYPE_INFO)
            Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
            Exit Sub

        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre el.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 3 Then
            Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        Select Case Npclist(.flags.TargetNPC).NPCtype

            Case eNPCType.Banquero
                Call WriteChatOverHead(UserIndex, "Tienes " & .Stats.Banco & " monedas de oro en tu cuenta.", Npclist(.flags.TargetNPC).Char.CharIndex, 255, 255, 255)
            
            Case eNPCType.Timbero

                If Not .flags.Privilegios And PlayerType.User Then
                    earnings = Apuestas.Ganancias - Apuestas.Perdidas
                    
                    If earnings >= 0 And Apuestas.Ganancias <> 0 Then
                        Percentage = Int(earnings * 100 / Apuestas.Ganancias)

                    End If
                    
                    If earnings < 0 And Apuestas.Perdidas <> 0 Then
                        Percentage = Int(earnings * 100 / Apuestas.Perdidas)

                    End If
                    
                    Call WriteConsoleMsg(UserIndex, "Entradas: " & Apuestas.Ganancias & " Salida: " & Apuestas.Perdidas & " Ganancia Neta: " & earnings & " (" & Percentage & "%) Jugadas: " & Apuestas.Jugadas, FontTypeNames.FONTTYPE_INFO)

                End If

        End Select

    End With

End Sub

''
' Handles the "PetStand" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePetStand(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead people can't use pets
        If .flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "Estas muerto!!", FontTypeNames.FONTTYPE_INFO)
            Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
            Exit Sub

        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre el.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        'Make sure it's close enough
        If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        'Make sure it's his pet
        If Npclist(.flags.TargetNPC).MaestroUser <> UserIndex Then Exit Sub
        
        'Do it!
        Npclist(.flags.TargetNPC).Movement = TipoAI.ESTATICO
        
        Call Expresar(.flags.TargetNPC, UserIndex)

    End With

End Sub

''
' Handles the "PetFollow" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePetFollow(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead users can't use pets
        If .flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "Estas muerto!!", FontTypeNames.FONTTYPE_INFO)
            Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
            Exit Sub

        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre el.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        'Make sure it's close enough
        If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        'Make usre it's the user's pet
        If Npclist(.flags.TargetNPC).MaestroUser <> UserIndex Then Exit Sub
        
        'Do it
        Call FollowAmo(.flags.TargetNPC)
        
        Call Expresar(.flags.TargetNPC, UserIndex)

    End With

End Sub

''
' Handles the "ReleasePet" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleReleasePet(ByVal UserIndex As Integer)

    '***************************************************
    'Author: ZaMa
    'Last Modification: 18/11/2009
    '
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead users can't use pets
        If .flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "Estas muerto!!", FontTypeNames.FONTTYPE_INFO)
            Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
            Exit Sub

        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar una mascota, haz click izquierdo sobre ella.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        'Make usre it's the user's pet
        If Npclist(.flags.TargetNPC).MaestroUser <> UserIndex Then Exit Sub
        
        'Make sure it's close enough
        If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        'Do it
        Call QuitarPet(UserIndex, .flags.TargetNPC)
            
    End With

End Sub

''
' Handles the "TrainList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTrainList(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre el.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        Call AccionParaEntrenador(UserIndex)

    End With

End Sub

''
' Handles the "Rest" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRest(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead users can't use pets
        If .flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "Estas muerto!! Solo puedes usar items cuando estas vivo.", FontTypeNames.FONTTYPE_INFO)
            Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
            Exit Sub

        End If
        
        If HayOBJarea(.Pos, FOGATA) Then
            Call WriteRestOK(UserIndex)
            
            If Not .flags.Descansar Then
                Call WriteConsoleMsg(UserIndex, "Te acomodas junto a la fogata y comienzas a descansar.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "Te levantas.", FontTypeNames.FONTTYPE_INFO)

            End If
            
            .flags.Descansar = Not .flags.Descansar
        Else

            If .flags.Descansar Then
                Call WriteRestOK(UserIndex)
                Call WriteConsoleMsg(UserIndex, "Te levantas.", FontTypeNames.FONTTYPE_INFO)
                
                .flags.Descansar = False
                Exit Sub

            End If
            
            Call WriteConsoleMsg(UserIndex, "No hay ninguna fogata junto a la cual descansar.", FontTypeNames.FONTTYPE_INFO)

        End If

    End With

End Sub

''
' Handles the "Meditate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleMeditate(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 04/15/08 (NicoNZ)
    'Arregle un bug que mandaba un index de la meditacion diferente
    'al que decia el server.
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead users can't use pets
        If .flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "Estas muerto!! Solo puedes meditar cuando estas vivo.", FontTypeNames.FONTTYPE_INFO)
            Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
            Exit Sub

        End If
        
        If .flags.Equitando Then
            Call WriteConsoleMsg(UserIndex, "No puedes meditar mientras si estas montado.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Can he meditate?
        If .Stats.MaxMAN = 0 Then
            Call WriteConsoleMsg(UserIndex, "Solo las clases magicas conocen el arte de la meditacion.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        'Admins don't have to wait :D
        If Not .flags.Privilegios And PlayerType.User Then
            .Stats.MinMAN = .Stats.MaxMAN
            Call WriteConsoleMsg(UserIndex, "Mana restaurado.", FontTypeNames.FONTTYPE_VENENO)
            Call WriteUpdateMana(UserIndex)
            Exit Sub

        End If
        
        Call WriteMeditateToggle(UserIndex)
        
        If .flags.Meditando Then Call WriteConsoleMsg(UserIndex, "Dejas de meditar.", FontTypeNames.FONTTYPE_INFO)
        
        .flags.Meditando = Not .flags.Meditando
        
        'Barrin 3/10/03 Tiempo de inicio al meditar
        If .flags.Meditando Then
            .Counters.tInicioMeditar = GetTickCount() And &H7FFFFFFF
            
            Call WriteConsoleMsg(UserIndex, "Te estas concentrando. En " & Fix(TIEMPO_INICIOMEDITAR / 1000) & " segundos comenzaras a meditar.", FontTypeNames.FONTTYPE_INFO)
            
            .Char.loops = INFINITE_LOOPS
            
            'Show proper FX according to level
            If .Stats.ELV < 13 Then
                .Char.FX = FXIDs.FXMEDITARCHICO
            
            ElseIf .Stats.ELV < 25 Then
                .Char.FX = FXIDs.FXMEDITARMEDIANO
            
            ElseIf .Stats.ELV < 35 Then
                .Char.FX = FXIDs.FXMEDITARGRANDE
            
            ElseIf .Stats.ELV < 42 Then
                .Char.FX = FXIDs.FXMEDITARXGRANDE
            
            Else
                .Char.FX = FXIDs.FXMEDITARXXGRANDE

            End If
            
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, .Char.FX, INFINITE_LOOPS))
        Else
            .Counters.bPuedeMeditar = False
            
            .Char.FX = 0
            .Char.loops = 0
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, 0, 0))

        End If

    End With

End Sub

''
' Handles the "Resucitate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleResucitate(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 07/01/20
    'Arreglo validacion de NPC para que funcione el comando. (Recox)
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Se asegura que el target es un npc
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre el.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        'Validate NPC and make sure player is dead
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Revividor Or .flags.Muerto = 0 Then Exit Sub
        
        'Make sure it's close enough
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 5 Then
            Call WriteConsoleMsg(UserIndex, "El sacerdote no puede resucitarte debido a que estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If

        Call SacerdoteResucitateUser(UserIndex)
    End With

End Sub

''
' Handles the "Consultation" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleConsultation(ByVal UserIndex As String)
    '***************************************************
    'Author: ZaMa
    'Last Modification: 01/05/2010
    'Habilita/Deshabilita el modo consulta.
    '01/05/2010: ZaMa - Agrego validaciones.
    '16/09/2010: ZaMa - No se hace visible en los clientes si estaba navegando (porque ya lo estaba).
    '***************************************************
    
    Dim UserConsulta As Integer
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        ' Comando exclusivo para gms
        If Not EsGm(UserIndex) Then Exit Sub
        
        UserConsulta = .flags.TargetUser
        
        'Se asegura que el target es un usuario
        If UserConsulta = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un usuario, haz click izquierdo sobre el.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        ' No podes ponerte a vos mismo en modo consulta.
        If UserConsulta = UserIndex Then Exit Sub
        
        ' No podes estra en consulta con otro gm
        If EsGm(UserConsulta) Then
            Call WriteConsoleMsg(UserIndex, "No puedes iniciar el modo consulta con otro administrador.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        Dim username As String

        username = UserList(UserConsulta).Name
        
        ' Si ya estaba en consulta, termina la consulta
        If UserList(UserConsulta).flags.EnConsulta Then
            Call WriteConsoleMsg(UserIndex, "Has terminado el modo consulta con " & username & ".", FontTypeNames.FONTTYPE_INFOBOLD)
            Call WriteConsoleMsg(UserConsulta, "Has terminado el modo consulta.", FontTypeNames.FONTTYPE_INFOBOLD)
            Call LogGM(.Name, "Termino consulta con " & username)
            
            UserList(UserConsulta).flags.EnConsulta = False
        
            ' Sino la inicia
        Else
            Call WriteConsoleMsg(UserIndex, "Has iniciado el modo consulta con " & username & ".", FontTypeNames.FONTTYPE_INFOBOLD)
            Call WriteConsoleMsg(UserConsulta, "Has iniciado el modo consulta.", FontTypeNames.FONTTYPE_INFOBOLD)
            Call LogGM(.Name, "Inicio consulta con " & username)
            
            With UserList(UserConsulta)
                .flags.EnConsulta = True
                
                ' Pierde invi u ocu
                If .flags.invisible = 1 Or .flags.Oculto = 1 Then
                    .flags.Oculto = 0
                    .flags.invisible = 0
                    .Counters.TiempoOculto = 0
                    .Counters.Invisibilidad = 0
                    
                    If UserList(UserConsulta).flags.Navegando = 0 Then
                        Call UsUaRiOs.SetInvisible(UserConsulta, UserList(UserConsulta).Char.CharIndex, False)

                    End If

                End If

            End With

        End If
        
        Call UsUaRiOs.SetConsulatMode(UserConsulta)

    End With

End Sub

''
' Handles the "RecordDetailsRequest" message.
'
' @param UserIndex The index of the user sending the message.
            
Public Sub HandleRecordDetailsRequest(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Amraphen
    'Last Modification: 07/04/2011
    'Handles the "RecordListRequest" message
    '***************************************************
    Dim RecordIndex As Byte

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        RecordIndex = .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster) Then Exit Sub
        
        Call WriteRecordDetails(UserIndex, RecordIndex)

    End With

End Sub

Public Sub HandleMoveItem(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Ignacio Mariano Tirabasso (Budi)
    'Last Modification: 01/01/2011
    '
    '***************************************************

    With UserList(UserIndex)

        Dim originalSlot As Byte

        Dim newSlot      As Byte
    
        Call .incomingData.ReadByte
    
        originalSlot = .incomingData.ReadByte
        newSlot = .incomingData.ReadByte
        Call .incomingData.ReadByte
    
        Call InvUsuario.moveItem(UserIndex, originalSlot, newSlot)
    
    End With

End Sub

''
' Handles the "LoginExistingAccount" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleLoginExistingAccount(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 12/10/2018
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 14 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errHandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue
    Set Buffer = New clsByteQueue

    Call Buffer.CopyBuffer(UserList(UserIndex).incomingData)
    
    'Remove packet ID
    Call Buffer.ReadByte

    Dim username    As String

    Dim Password    As String

    Dim version     As String
    
    Dim macAddress  As String

    Dim hdSerial    As Long
    
    username = Buffer.ReadASCIIString()
    Password = Buffer.ReadASCIIString()

    'Convert version number to string
    version = CStr(Buffer.ReadByte()) & "." & CStr(Buffer.ReadByte()) & "." & CStr(Buffer.ReadByte())
    
    'Seguridad LwK
    macAddress = Buffer.ReadASCIIString()
    hdSerial = Buffer.ReadLong()
    
    If Not VersionOK(version) Then
        Call WriteErrorMsg(UserIndex, "Esta version del juego es obsoleta, la version correcta es la " & ULTIMAVERSION & ". La misma se encuentra disponible en http://winterao.com.ar")
        
    Else
        Call ConnectAccount(UserIndex, username, Password, macAddress, hdSerial)

    End If

    'If we got here then packet is complete, copy data back to original queue
    Call UserList(UserIndex).incomingData.CopyBuffer(Buffer)
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "Heal" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleHeal(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Se asegura que el target es un npc
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre el.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        If (Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Revividor) Or .flags.Muerto <> 0 Then Exit Sub
        
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "El sacerdote no puede curarte debido a que estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        Call SacerdoteHealUser(UserIndex)
    End With

End Sub

''
' Handles the "RequestStats" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestStats(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    Call SendUserStatsTxt(UserIndex, UserIndex)

End Sub

''
' Handles the "Help" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleHelp(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    Call SendHelp(UserIndex)

End Sub

''
' Handles the "CommerceStart" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCommerceStart(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    Dim i As Integer

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead people can't commerce
        If .flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "Estas muerto!!", FontTypeNames.FONTTYPE_INFO)
            Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
            Exit Sub

        End If
        
        'Is it already in commerce mode??
        If .flags.Comerciando Then
            Call WriteConsoleMsg(UserIndex, "Ya estas comerciando.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        'Validate target NPC
        If .flags.TargetNPC > 0 Then

            'Does the NPC want to trade??
            If Npclist(.flags.TargetNPC).Comercia = 0 Then

                If LenB(Npclist(.flags.TargetNPC).Desc) <> 0 Then
                    Call WriteChatOverHead(UserIndex, "No tengo ningun interes en comerciar.", Npclist(.flags.TargetNPC).Char.CharIndex, 255, 255, 255)
                End If
                
                Exit Sub

            End If
            
            If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 3 Then
                Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
            
            'Start commerce....
            Call IniciarComercioNPC(UserIndex)
            '[Alejo]
        ElseIf .flags.TargetUser > 0 Then

            'User commerce...
            'Can he commerce??
            If .flags.Privilegios And PlayerType.Consejero Then
                Call WriteConsoleMsg(UserIndex, "No puedes vender items.", FontTypeNames.FONTTYPE_WARNING)
                Exit Sub

            End If
            
            'Is the other one dead??
            If UserList(.flags.TargetUser).flags.Muerto = 1 Then
                Call WriteConsoleMsg(UserIndex, "No puedes comerciar con los muertos!!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
            
            'Is it me??
            If .flags.TargetUser = UserIndex Then
                Call WriteConsoleMsg(UserIndex, "No puedes comerciar con vos mismo!!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
            
            'Check distance
            If Distancia(UserList(.flags.TargetUser).Pos, .Pos) > 3 Then
                Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos del usuario.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
            
            'Is he already trading?? is it with me or someone else??
            If UserList(.flags.TargetUser).flags.Comerciando = True And UserList(.flags.TargetUser).ComUsu.DestUsu <> UserIndex Then
                Call WriteConsoleMsg(UserIndex, "No puedes comerciar con el usuario en este momento.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
            
            'Initialize some variables...
            .ComUsu.DestUsu = .flags.TargetUser
            .ComUsu.DestNick = UserList(.flags.TargetUser).Name

            For i = 1 To MAX_OFFER_SLOTS
                .ComUsu.cant(i) = 0
                .ComUsu.Objeto(i) = 0
            Next i

            .ComUsu.GoldAmount = 0
            
            .ComUsu.Acepto = False
            .ComUsu.Confirmo = False
            
            'Rutina para comerciar con otro usuario
            Call IniciarComercioConUsuario(UserIndex, .flags.TargetUser)
        Else
            Call WriteConsoleMsg(UserIndex, "Primero haz click izquierdo sobre el personaje.", FontTypeNames.FONTTYPE_INFO)

        End If

    End With

End Sub

''
' Handles the "BankStart" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankStart(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead people can't commerce
        If .flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "Estas muerto!!", FontTypeNames.FONTTYPE_INFO)
            Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
            Exit Sub

        End If
        
        If .flags.Comerciando Then
            Call WriteConsoleMsg(UserIndex, "Ya estas comerciando.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        'Validate target NPC
        If .flags.TargetNPC > 0 Then
            If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 3 Then
                Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos del vendedor.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
            
            'If it's the banker....
            If Npclist(.flags.TargetNPC).NPCtype = eNPCType.Banquero Then
                Call IniciarDeposito(UserIndex)

            End If

        Else
            Call WriteConsoleMsg(UserIndex, "Primero haz click izquierdo sobre el personaje.", FontTypeNames.FONTTYPE_INFO)

        End If

    End With

End Sub

''
' Handles the "Enlist" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleEnlist(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre el.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Noble Or .flags.Muerto <> 0 Then Exit Sub
        
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 4 Then
            Call WriteConsoleMsg(UserIndex, "Debes acercarte mas.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        If Npclist(.flags.TargetNPC).flags.Faccion = 0 Then
            Call EnlistarArmadaReal(UserIndex)
        Else
            Call EnlistarCaos(UserIndex)

        End If

    End With

End Sub

''
' Handles the "Information" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleInformation(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    Dim Matados    As Integer

    Dim NextRecom  As Integer

    Dim Diferencia As Integer
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre el.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Noble Or .flags.Muerto <> 0 Then Exit Sub
        
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 4 Then
            Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        NextRecom = .Faccion.NextRecompensa
        
        If Npclist(.flags.TargetNPC).flags.Faccion = 0 Then
            If .Faccion.ArmadaReal = 0 Then
                Call WriteChatOverHead(UserIndex, "No perteneces a las tropas reales!!", Npclist(.flags.TargetNPC).Char.CharIndex, 255, 255, 255)
                Exit Sub

            End If
            
            Matados = .Faccion.CriminalesMatados
            Diferencia = NextRecom - Matados
            
            If Diferencia > 0 Then
                Call WriteChatOverHead(UserIndex, "Tu deber es combatir criminales, mata " & Diferencia & " criminales mas y te dare una recompensa.", Npclist(.flags.TargetNPC).Char.CharIndex, 255, 255, 255)
            Else
                Call WriteChatOverHead(UserIndex, "Tu deber es combatir criminales, y ya has matado los suficientes como para merecerte una recompensa.", Npclist(.flags.TargetNPC).Char.CharIndex, 255, 255, 255)

            End If

        Else

            If .Faccion.FuerzasCaos = 0 Then
                Call WriteChatOverHead(UserIndex, "No perteneces a la legion oscura!!", Npclist(.flags.TargetNPC).Char.CharIndex, 255, 255, 255)
                Exit Sub

            End If
            
            Matados = .Faccion.CiudadanosMatados
            Diferencia = NextRecom - Matados
            
            If Diferencia > 0 Then
                Call WriteChatOverHead(UserIndex, "Tu deber es sembrar el caos y la desesperanza, mata " & Diferencia & " ciudadanos mas y te dare una recompensa.", Npclist(.flags.TargetNPC).Char.CharIndex, 255, 255, 255)
            Else
                Call WriteChatOverHead(UserIndex, "Tu deber es sembrar el caos y la desesperanza, y creo que estas en condiciones de merecer una recompensa.", Npclist(.flags.TargetNPC).Char.CharIndex, 255, 255, 255)

            End If

        End If

    End With

End Sub

''
' Handles the "Reward" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleReward(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre el.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Noble Or .flags.Muerto <> 0 Then Exit Sub
        
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 4 Then
            Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        If Npclist(.flags.TargetNPC).flags.Faccion = 0 Then
            If .Faccion.ArmadaReal = 0 Then
                Call WriteChatOverHead(UserIndex, "No perteneces a las tropas reales!!", Npclist(.flags.TargetNPC).Char.CharIndex, 255, 255, 255)
                Exit Sub

            End If

            Call RecompensaArmadaReal(UserIndex)
        Else

            If .Faccion.FuerzasCaos = 0 Then
                Call WriteChatOverHead(UserIndex, "No perteneces a la legion oscura!!", Npclist(.flags.TargetNPC).Char.CharIndex, 255, 255, 255)
                Exit Sub

            End If

            Call RecompensaCaos(UserIndex)

        End If

    End With

End Sub

''
' Handles the "RequestMOTD" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestMOTD(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    Call SendMOTD(UserIndex)

End Sub

''
' Handles the "UpTime" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUpTime(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 01/10/08
    '01/10/2008 - Marcos Martinez (ByVal) - Automatic restart removed from the server along with all their assignments and varibles
    '***************************************************
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    Dim time      As Long

    Dim UpTimeStr As String
    
    Call WriteConsoleServerUpTimeMsg(UserIndex)
End Sub

''
' Handles the "PartyLeave" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePartyLeave(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    Call mdParty.SalirDeParty(UserIndex)

End Sub

''
' Handles the "ShareNpc" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleShareNpc(ByVal UserIndex As Integer)
    '***************************************************
    'Author: ZaMa
    'Last Modification: 15/04/2010
    'Shares owned npcs with other user
    '***************************************************
    
    Dim TargetUserIndex  As Integer

    Dim SharingUserIndex As Integer
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        ' Didn't target any user
        TargetUserIndex = .flags.TargetUser

        If TargetUserIndex = 0 Then Exit Sub
        
        ' Can't share with admins
        If EsGm(TargetUserIndex) Then
            Call WriteConsoleMsg(UserIndex, "No puedes compartir npcs con administradores!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        ' Pk or Caos?
        If criminal(UserIndex) Then

            ' Caos can only share with other caos
            If esCaos(UserIndex) Then
                If Not esCaos(TargetUserIndex) Then
                    Call WriteConsoleMsg(UserIndex, "Solo puedes compartir npcs con miembros de tu misma faccion!!", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub

                End If
                
                ' Pks don't need to share with anyone
            Else
                Exit Sub

            End If
        
            ' Ciuda or Army?
        Else

            ' Can't share
            If criminal(TargetUserIndex) Then
                Call WriteConsoleMsg(UserIndex, "No puedes compartir npcs con criminales!!", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

        End If
        
        ' Already sharing with target
        SharingUserIndex = .flags.ShareNpcWith

        If SharingUserIndex = TargetUserIndex Then Exit Sub
        
        ' Aviso al usuario anterior que dejo de compartir
        If SharingUserIndex <> 0 Then
            Call WriteConsoleMsg(SharingUserIndex, .Name & " ha dejado de compartir sus npcs contigo.", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(UserIndex, "Has dejado de compartir tus npcs con " & UserList(SharingUserIndex).Name & ".", FontTypeNames.FONTTYPE_INFO)

        End If
        
        .flags.ShareNpcWith = TargetUserIndex
        
        Call WriteConsoleMsg(TargetUserIndex, .Name & " ahora comparte sus npcs contigo.", FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(UserIndex, "Ahora compartes tus npcs con " & UserList(TargetUserIndex).Name & ".", FontTypeNames.FONTTYPE_INFO)
        
    End With
    
End Sub

''
' Handles the "StopSharingNpc" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleStopSharingNpc(ByVal UserIndex As Integer)
    '***************************************************
    'Author: ZaMa
    'Last Modification: 15/04/2010
    'Stop Sharing owned npcs with other user
    '***************************************************
    
    Dim SharingUserIndex As Integer
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        SharingUserIndex = .flags.ShareNpcWith
        
        If SharingUserIndex <> 0 Then
            
            ' Aviso al que compartia y al que le compartia.
            Call WriteConsoleMsg(SharingUserIndex, .Name & " ha dejado de compartir sus npcs contigo.", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(SharingUserIndex, "Has dejado de compartir tus npcs con " & UserList(SharingUserIndex).Name & ".", FontTypeNames.FONTTYPE_INFO)
            
            .flags.ShareNpcWith = 0

        End If
        
    End With

End Sub

''
' Handles the "Inquiry" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleInquiry(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    ConsultaPopular.SendInfoEncuesta (UserIndex)

End Sub

''
' Handles the "GuildMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildMessage(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 15/07/2009
    '02/03/2009: ZaMa - Arreglado un indice mal pasado a la funcion de cartel de clanes overhead.
    '15/07/2009: ZaMa - Now invisible admins only speak by console
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim Chat As String
        
        Chat = Buffer.ReadASCIIString()
        
        If LenB(Chat) <> 0 Then
            'Analize chat...
            Call Statistics.ParseChat(Chat)
            
            If .GuildIndex > 0 Then
                Call SendData(SendTarget.ToDiosesYclan, .GuildIndex, PrepareMessageGuildChat(.Name & "> " & Chat))
                
                If Not (.flags.AdminInvisible = 1) Then Call SendData(SendTarget.ToClanArea, UserIndex, PrepareMessageChatOverHead("< " & Chat & " >", .Char.CharIndex, 255, 255, 0))

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "PartyMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePartyMessage(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim Chat As String
        
        Chat = Buffer.ReadASCIIString()
        
        If LenB(Chat) <> 0 Then
            'Analize chat...
            Call Statistics.ParseChat(Chat)
            
            Call mdParty.BroadCastParty(UserIndex, Chat)

            'TODO : Con la 0.12.1 se debe definir si esto vuelve o se borra (/CMSG overhead)
            'Call SendData(SendTarget.ToPartyArea, UserIndex, UserList(UserIndex).Pos.map, "||" & vbYellow & "Â°< " & mid$(rData, 7) & " >Â°" & CStr(UserList(UserIndex).Char.CharIndex))
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

Public Sub HandleSearchNpc(ByVal UserIndex As Integer)
 
    On Error GoTo errHandler

    With UserList(UserIndex)

        Dim Buffer As New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
       
        Call Buffer.ReadByte
       
        Dim i       As Long

        Dim n       As Integer

        Dim Name    As String

        Dim UserNpc As String

        Dim tStr    As String

        UserNpc = Buffer.ReadASCIIString()
        
        Call .incomingData.CopyBuffer(Buffer)
        
        ' Es Game-Master?
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        tStr = Tilde(UserNpc)
      
        For i = 1 To val(LeerNPCs.GetValue("INIT", "NumNPCs"))
            Name = LeerNPCs.GetValue("NPC" & i, "Name")
       
            If InStr(1, Tilde(Name), tStr) Then
                Call WriteSearchList(UserIndex, i, CStr(i & " - " & Name), False)
                n = n + 1

            End If

        Next i
   
        If n = 0 Then
            Call WriteSearchList(UserIndex, 0, "No hubo resultados de la busqueda.", False)

        End If

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
   
    Set Buffer = Nothing
   
    If Error <> 0 Then Err.Raise Error

End Sub
 
Private Sub HandleSearchObj(ByVal UserIndex As Integer)
       
    On Error GoTo errHandler

    With UserList(UserIndex)

        Dim Buffer As New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
           
        Call Buffer.ReadByte
           
        Dim UserObj As String

        Dim tUser   As Integer

        Dim n       As Integer

        Dim i       As Long

        Dim tStr    As String
       
        UserObj = Buffer.ReadASCIIString()
        
        Call .incomingData.CopyBuffer(Buffer)
        
        ' Es Game-Master?
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub

        tStr = Tilde(UserObj)
          
        For i = 1 To UBound(ObjData)

            If InStr(1, Tilde(ObjData(i).Name), tStr) Then
                Call WriteSearchList(UserIndex, i, CStr(i & " - " & ObjData(i).Name), True)
                n = n + 1

            End If

        Next

        If n = 0 Then
            Call WriteSearchList(UserIndex, 0, "No hubo resultados de la busqueda.", False)

        End If
                
    End With
     
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
       
    Set Buffer = Nothing
       
    If Error <> 0 Then Err.Raise Error
        
End Sub

Private Sub HandleEnviaCvc(ByVal UserIndex As Integer)

    'Dim targetIndex As Integer

    With UserList(UserIndex)
        .incomingData.ReadByte

        If .flags.TargetUser = 0 Then Exit Sub 'gdk: adonde mierda clickeas manko?
        Call Mod_ClanvsClan.Enviar(UserIndex, .flags.TargetUser)

    End With

End Sub

Private Sub HandleAceptarCvc(ByVal UserIndex As Integer)

    With UserList(UserIndex)
        .incomingData.ReadByte

        If .flags.TargetUser = 0 Then Exit Sub
        Call Mod_ClanvsClan.Aceptar(UserIndex, .flags.TargetUser)

    End With

End Sub

Private Sub HandleIrCvc(ByVal UserIndex As Integer)

    With UserList(UserIndex)
        .incomingData.ReadByte
                
        Call Mod_ClanvsClan.ConectarCVC(UserIndex, True)  'gdk: si le pones false bugeas toditus.

    End With

End Sub

Public Sub HandleDragAndDropHechizos(ByVal UserIndex As Integer)
 
    With UserList(UserIndex)
        
        Call .incomingData.ReadByte
        
        Dim AnteriorPosicion As Integer: AnteriorPosicion = .incomingData.ReadInteger
        Dim NuevaPosicion As Integer: NuevaPosicion = .incomingData.ReadInteger
        
        Dim Hechizo As Integer: Hechizo = .Stats.UserHechizos(NuevaPosicion)

        .Stats.UserHechizos(NuevaPosicion) = .Stats.UserHechizos(AnteriorPosicion)
        .Stats.UserHechizos(AnteriorPosicion) = Hechizo
             
    End With

End Sub

''
' Handles the "GuildOnline" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildOnline(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim onlineList As String
        
        onlineList = modGuilds.m_ListaDeMiembrosOnline(UserIndex, .GuildIndex)
        
        If .GuildIndex <> 0 Then
            Call WriteConsoleMsg(UserIndex, "Companeros de tu clan conectados: " & onlineList, FontTypeNames.FONTTYPE_GUILDMSG)
        Else
            Call WriteConsoleMsg(UserIndex, "No pertences a ningUn clan.", FontTypeNames.FONTTYPE_GUILDMSG)

        End If

    End With

End Sub

''
' Handles the "PartyOnline" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePartyOnline(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    Call mdParty.OnlineParty(UserIndex)

End Sub

''
' Handles the "CouncilMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCouncilMessage(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim Chat As String
        
        Chat = Buffer.ReadASCIIString()
        
        If LenB(Chat) <> 0 Then
            'Analize chat...
            Call Statistics.ParseChat(Chat)
            
            If .flags.Privilegios And PlayerType.RoyalCouncil Then
                Call SendData(SendTarget.ToConsejo, UserIndex, PrepareMessageConsoleMsg("(Consejero) " & .Name & "> " & Chat, FontTypeNames.FONTTYPE_CONSEJO))
            ElseIf .flags.Privilegios And PlayerType.ChaosCouncil Then
                Call SendData(SendTarget.ToConsejoCaos, UserIndex, PrepareMessageConsoleMsg("(Consejero) " & .Name & "> " & Chat, FontTypeNames.FONTTYPE_CONSEJOCAOS))

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "RoleMasterRequest" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRoleMasterRequest(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim request As String
        
        request = Buffer.ReadASCIIString()
        
        If LenB(request) <> 0 Then
            Call WriteConsoleMsg(UserIndex, "Su solicitud ha sido enviada.", FontTypeNames.FONTTYPE_INFO)
            Call SendData(SendTarget.ToRolesMasters, 0, PrepareMessageConsoleMsg(.Name & " PREGUNTA ROL: " & request, FontTypeNames.FONTTYPE_GUILDMSG))

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "GMRequest" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGMRequest(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'
'***************************************************
    Dim Tipo As Byte
    Dim Message As String
    'Bug y Sugerencias
    Dim cant As Integer
    Dim Motivo As Integer
    Dim Nuevo As String
    Dim Mensaje As String
    Dim FileDir As String

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Tipo = .incomingData.ReadByte
        Message = .incomingData.ReadASCIIString()
        
        'Ruta donde se guardan los reportes
        FileDir = App.Path & "\logs\REPORTES\"
        
        'Si es una Consulta:
        Select Case Tipo
        
        Case 0 'Consultas
        
            If Not Ayuda.Existe(.Name) Then
                Call WriteConsoleMsg(UserIndex, "El mensaje ha sido entregado, ahora sólo debes esperar que se desocupe algún GM.", FontTypeNames.FONTTYPE_INFO)
                Call Ayuda.Push(.Name & ";" & Message)
                Exit Sub
            Else
                Call Ayuda.Quitar(.Name)
                Call Ayuda.Push(.Name & ";" & Message)
                Call WriteConsoleMsg(UserIndex, "Ya habías mandado un mensaje, tu mensaje ha sido movido al final de la cola de mensajes.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
        Case 1 'Reporte de bugs
            
            If FileExist(FileDir, vbDirectory) = False Then _
                MkDir FileDir
            
            cant = GetVar(FileDir & "Bugs.INI", "BUGS", "CANTIDAD")
            Motivo = val(cant) + 1
            Nuevo = "Bug" & Motivo
            Mensaje = Date & " " & time & " - " & UserList(UserIndex).Name & " Reporto el siguiente Bug: " & Message & " - IP: " & UserList(UserIndex).IP

            Call WriteVar(FileDir & "Bugs.INI", "Bugs", "Cantidad", Motivo)
            Call WriteVar(FileDir & "Bugs.INI", "Reportes", Nuevo, Mensaje)
            
            Call WriteConsoleMsg(UserIndex, "El Bug ha sido reportado exitosamente! Gracias por colaborar con WinterAO.", FONTTYPE_GUILD)
            Call WriteConsoleMsg(SendTarget.ToAdmins, Mensaje, FONTTYPE_TALK)
            
        Case 2 'Sugerencia
            
            If FileExist(FileDir, vbDirectory) = False Then _
                MkDir FileDir
        
            cant = GetVar(FileDir & "Sugerencias.ini", "SUGERENCIAS", "CANTIDAD")
            Motivo = val(cant) + 1
            Nuevo = "Sugerencia" & Motivo
            Mensaje = Date & " " & time & " - " & UserList(UserIndex).Name & " Reporto la siguiente sugerencia: " & Message & " - IP: " & UserList(UserIndex).IP

            Call WriteVar(FileDir & "Sugerencias.ini", "SUGERENCIAS", "Cantidad", Motivo)
            Call WriteVar(FileDir & "Sugerencias.ini", "Reportes", Nuevo, Mensaje)
            
            Call WriteConsoleMsg(UserIndex, "La sugerencia ha sido guardada! Gracias por colaboar con WinterAO.", FONTTYPE_GUILD)
            Call WriteConsoleMsg(SendTarget.ToAdmins, Mensaje, FONTTYPE_TALK)
            
        Case 3 'Denuncia
            
            If FileExist(FileDir, vbDirectory) = False Then _
                MkDir FileDir
        
            cant = GetVar(FileDir & "Sugerencias.ini", "DENUNCIAS", "CANTIDAD")
            Motivo = val(cant) + 1
            Nuevo = "Sugerencia" & Motivo
            Mensaje = Date & " " & time & " - " & UserList(UserIndex).Name & " Reporto la siguiente denunciaa: " & Message & " - IP: " & UserList(UserIndex).IP

            Call WriteVar(FileDir & "Denuncias.ini", "SUGERENCIAS", "Cantidad", Motivo)
            Call WriteVar(FileDir & "Denuncias.ini", "Reportes", Nuevo, Mensaje)
            
            Call WriteConsoleMsg(UserIndex, "La denuncia ha sido registrada! Gracias por colaboar con WinterAO.", FONTTYPE_GUILD)
            Call WriteConsoleMsg(SendTarget.ToAdmins, Mensaje, FONTTYPE_TALK)
            
        End Select
        
    End With
End Sub

''
' Handles the "ChangeDescription" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleChangeDescription(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim description As String
        
        description = Buffer.ReadASCIIString()

        If Not AsciiValidos(description) Then
            Call WriteConsoleMsg(UserIndex, "La descripcion tiene caracteres invalidos.", FontTypeNames.FONTTYPE_INFO)
        Else
            .Desc = Trim$(description)
            Call WriteConsoleMsg(UserIndex, "La descripcion ha cambiado.", FontTypeNames.FONTTYPE_INFO)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "GuildVote" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildVote(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim vote     As String

        Dim errorStr As String
        
        vote = Buffer.ReadASCIIString()
        
        If Not modGuilds.v_UsuarioVota(UserIndex, vote, errorStr) Then
            Call WriteConsoleMsg(UserIndex, "Voto NO contabilizado: " & errorStr, FontTypeNames.FONTTYPE_GUILD)
        Else
            Call WriteConsoleMsg(UserIndex, "Voto contabilizado.", FontTypeNames.FONTTYPE_GUILD)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "ShowGuildNews" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleShowGuildNews(ByVal UserIndex As Integer)
    '***************************************************
    'Author: ZaMA
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    With UserList(UserIndex)
        
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Call modGuilds.SendGuildNews(UserIndex)

    End With

End Sub

''
' Handles the "Punishments" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePunishments(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 25/08/2009
    '25/08/2009: ZaMa - Now only admins can see other admins' punishment list
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim Name  As String

        Dim Count As Integer
        
        Name = Buffer.ReadASCIIString()
        
        If LenB(Name) <> 0 Then
            If (InStrB(Name, "\") <> 0) Then
                Name = Replace(Name, "\", "")

            End If

            If (InStrB(Name, "/") <> 0) Then
                Name = Replace(Name, "/", "")

            End If

            If (InStrB(Name, ":") <> 0) Then
                Name = Replace(Name, ":", "")

            End If

            If (InStrB(Name, "|") <> 0) Then
                Name = Replace(Name, "|", "")

            End If
            
            If (EsAdmin(Name) Or EsDios(Name) Or EsSemiDios(Name) Or EsConsejero(Name) Or EsRolesMaster(Name)) And (UserList(UserIndex).flags.Privilegios And PlayerType.User) Then
                Call WriteConsoleMsg(UserIndex, "No puedes ver las penas de los administradores.", FontTypeNames.FONTTYPE_INFO)
            Else

                If PersonajeExiste(Name) Then
                    Count = GetUserAmountOfPunishments(Name)

                    If Count = 0 Then
                        Call WriteConsoleMsg(UserIndex, "Sin prontuario..", FontTypeNames.FONTTYPE_INFO)
                    Else
                        Call SendUserPunishments(UserIndex, Name, Count)

                    End If

                Else
                    Call WriteConsoleMsg(UserIndex, "Personaje """ & Name & """ inexistente.", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "Gamble" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGamble(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '10/07/2010: ZaMa - Now normal npcs don't answer if asked to gamble.
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Amount  As Integer

        Dim TypeNpc As eNPCType
        
        Amount = .incomingData.ReadInteger()
        
        ' Dead?
        If .flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "Estas muerto!!", FontTypeNames.FONTTYPE_INFO)
            Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
        
            'Validate target NPC
        ElseIf .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre el.", FontTypeNames.FONTTYPE_INFO)
        
            ' Validate Distance
        ElseIf Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
        
            ' Validate NpcType
        ElseIf Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Timbero Then
            
            Dim TargetNpcType As eNPCType

            TargetNpcType = Npclist(.flags.TargetNPC).NPCtype
            
            ' Normal npcs don't speak
            If TargetNpcType <> eNPCType.Comun And TargetNpcType <> eNPCType.DRAGON And TargetNpcType <> eNPCType.Pretoriano Then
                Call WriteChatOverHead(UserIndex, "No tengo ningUn interes en apostar.", Npclist(.flags.TargetNPC).Char.CharIndex, 255, 255, 255)

            End If
            
            ' Validate amount
        ElseIf Amount < 1 Then
            Call WriteChatOverHead(UserIndex, "El minimo de apuesta es 1 moneda.", Npclist(.flags.TargetNPC).Char.CharIndex, 255, 255, 255)
        
            ' Validate amount
        ElseIf Amount > 5000 Then
            Call WriteChatOverHead(UserIndex, "El maximo de apuesta es 5000 monedas.", Npclist(.flags.TargetNPC).Char.CharIndex, 255, 255, 255)
        
            ' Validate user gold
        ElseIf .Stats.Gld < Amount Then
            Call WriteChatOverHead(UserIndex, "No tienes esa cantidad.", Npclist(.flags.TargetNPC).Char.CharIndex, 255, 255, 255)
        
        Else

            If RandomNumber(1, 100) <= 47 Then
                .Stats.Gld = .Stats.Gld + Amount
                Call WriteChatOverHead(UserIndex, "Felicidades! Has ganado " & CStr(Amount) & " monedas de oro.", Npclist(.flags.TargetNPC).Char.CharIndex, 255, 255, 255)
                
                Apuestas.Perdidas = Apuestas.Perdidas + Amount
                Call WriteVar(DatPath & "apuestas.dat", "Main", "Perdidas", CStr(Apuestas.Perdidas))
            Else
                .Stats.Gld = .Stats.Gld - Amount
                Call WriteChatOverHead(UserIndex, "Lo siento, has perdido " & CStr(Amount) & " monedas de oro.", Npclist(.flags.TargetNPC).Char.CharIndex, 255, 255, 255)
                
                Apuestas.Ganancias = Apuestas.Ganancias + Amount
                Call WriteVar(DatPath & "apuestas.dat", "Main", "Ganancias", CStr(Apuestas.Ganancias))

            End If
            
            Apuestas.Jugadas = Apuestas.Jugadas + 1
            
            Call WriteVar(DatPath & "apuestas.dat", "Main", "Jugadas", CStr(Apuestas.Jugadas))
            
            Call WriteUpdateGold(UserIndex)

        End If

    End With

End Sub

''
' Handles the "InquiryVote" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleInquiryVote(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim opt As Byte
        
        opt = .incomingData.ReadByte()
        
        Call WriteConsoleMsg(UserIndex, ConsultaPopular.doVotar(UserIndex, opt), FontTypeNames.FONTTYPE_GUILD)

    End With

End Sub

''
' Handles the "BankExtractGold" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankExtractGold(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Amount As Long
        
        Amount = .incomingData.ReadLong()
        
        'Dead people can't leave a faction.. they can't talk...
        If .flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "Estas muerto!!", FontTypeNames.FONTTYPE_INFO)
            Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
            Exit Sub

        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre el.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Banquero Then Exit Sub
        
        If Distancia(.Pos, Npclist(.flags.TargetNPC).Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        If Amount > 0 And Amount <= .Stats.Banco Then
            .Stats.Banco = .Stats.Banco - Amount
            .Stats.Gld = .Stats.Gld + Amount
            Call WriteChatOverHead(UserIndex, "Tenes " & .Stats.Banco & " monedas de oro en tu cuenta.", Npclist(.flags.TargetNPC).Char.CharIndex, 255, 255, 255)
        Else
            Call WriteChatOverHead(UserIndex, "No tienes esa cantidad.", Npclist(.flags.TargetNPC).Char.CharIndex, 255, 255, 255)

        End If
        
        Call WriteUpdateGold(UserIndex)
        Call WriteUpdateBankGold(UserIndex)

    End With

End Sub

''
' Handles the "LeaveFaction" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleLeaveFaction(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 09/28/2010
    ' 09/28/2010 C4b3z0n - Ahora la respuesta de los NPCs sino perteneces a ninguna faccion solo la hacen el Rey o el Demonio
    ' 05/17/06 - Maraxus
    '***************************************************

    Dim TalkToKing  As Boolean

    Dim TalkToDemon As Boolean

    Dim NPCIndex    As Integer
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        'Dead people can't leave a faction.. they can't talk...
        If .flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "Estas muerto!!", FontTypeNames.FONTTYPE_INFO)
            Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
            Exit Sub

        End If
        
        ' Chequea si habla con el rey o el demonio. Puede salir sin hacerlo, pero si lo hace le reponden los npcs
        NPCIndex = .flags.TargetNPC

        If NPCIndex <> 0 Then

            ' Es rey o domonio?
            If Npclist(NPCIndex).NPCtype = eNPCType.Noble Then

                'Rey?
                If Npclist(NPCIndex).flags.Faccion = 0 Then
                    TalkToKing = True
                    ' Demonio
                Else
                    TalkToDemon = True

                End If

            End If

        End If
               
        'Quit the Royal Army?
        If .Faccion.ArmadaReal = 1 Then

            ' Si le pidio al demonio salir de la armada, este le responde.
            If TalkToDemon Then
                Call WriteChatOverHead(UserIndex, "Sal de aqui bufon!!!", Npclist(NPCIndex).Char.CharIndex, 255, 255, 255)
            
            Else

                ' Si le pidio al rey salir de la armada, le responde.
                If TalkToKing Then
                    Call WriteChatOverHead(UserIndex, "Seras bienvenido a las fuerzas imperiales si deseas regresar.", Npclist(NPCIndex).Char.CharIndex, 255, 255, 255)

                End If
                
                Call ExpulsarFaccionReal(UserIndex, False)
                
            End If
        
            'Quit the Chaos Legion?
        ElseIf .Faccion.FuerzasCaos = 1 Then

            ' Si le pidio al rey salir del caos, le responde.
            If TalkToKing Then
                Call WriteChatOverHead(UserIndex, "Sal de aqui maldito criminal!!!", Npclist(NPCIndex).Char.CharIndex, 255, 255, 255)
            Else

                ' Si le pidio al demonio salir del caos, este le responde.
                If TalkToDemon Then
                    Call WriteChatOverHead(UserIndex, "Ya volveras arrastrandote.", Npclist(NPCIndex).Char.CharIndex, 255, 255, 255)

                End If
                
                Call ExpulsarFaccionCaos(UserIndex, False)

            End If

            ' No es faccionario
        Else
        
            ' Si le hablaba al rey o demonio, le repsonden ellos
            'Corregido, solo si son en efecto el rey o el demonio, no cualquier NPC (C4b3z0n)
            If (TalkToDemon And criminal(UserIndex)) Or (TalkToKing And Not criminal(UserIndex)) Then 'Si se pueden unir a la faccion (status), son invitados
                Call WriteChatOverHead(UserIndex, "No perteneces a nuestra faccion. Si deseas unirte, di /ENLISTAR", Npclist(NPCIndex).Char.CharIndex, 255, 255, 255)
            ElseIf (TalkToDemon And Not criminal(UserIndex)) Then
                Call WriteChatOverHead(UserIndex, "Sal de aqui bufon!!!", Npclist(NPCIndex).Char.CharIndex, 255, 255, 255)
            ElseIf (TalkToKing And criminal(UserIndex)) Then
                Call WriteChatOverHead(UserIndex, "Sal de aqui maldito criminal!!!", Npclist(NPCIndex).Char.CharIndex, 255, 255, 255)
            Else
                Call WriteConsoleMsg(UserIndex, "No perteneces a ninguna faccion!", FontTypeNames.FONTTYPE_FIGHT)

            End If
        
        End If
        
    End With
    
End Sub

''
' Handles the "BankDepositGold" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBankDepositGold(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Amount As Long
        
        Amount = .incomingData.ReadLong()
        
        'Dead people can't leave a faction.. they can't talk...
        If .flags.Muerto = 1 Then
            'Call WriteConsoleMsg(UserIndex, "Estas muerto!!", FontTypeNames.FONTTYPE_INFO)
            Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
            Exit Sub

        End If
        
        'Validate target NPC
        If .flags.TargetNPC = 0 Then
            Call WriteConsoleMsg(UserIndex, "Primero tienes que seleccionar un personaje, haz click izquierdo sobre el.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        If Distancia(Npclist(.flags.TargetNPC).Pos, .Pos) > 10 Then
            Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        If Npclist(.flags.TargetNPC).NPCtype <> eNPCType.Banquero Then Exit Sub
        
        'Calculamos la diferencia con el maximo de oro permitido el cual es el valor de LONG
        Dim RemainingAmountToMaximumGold As Long
        RemainingAmountToMaximumGold = 2147483647 - .Stats.Gld

        If .Stats.Banco >= 2147483647 And RemainingAmountToMaximumGold <= Amount Then
            Call WriteChatOverHead(UserIndex, "No puedes depositar el oro por que tendrias mas del maximo permitido (2147483647)", Npclist(.flags.TargetNPC).Char.CharIndex, 255, 0, 0)

        ElseIf Amount > 0 And Amount <= .Stats.Gld Then
            .Stats.Banco = .Stats.Banco + Amount
            .Stats.Gld = .Stats.Gld - Amount
            Call WriteChatOverHead(UserIndex, "Tenes " & .Stats.Banco & " monedas de oro en tu cuenta.", Npclist(.flags.TargetNPC).Char.CharIndex, 255, 255, 255)
            
            Call WriteUpdateGold(UserIndex)
            Call WriteUpdateBankGold(UserIndex)
        Else
            Call WriteChatOverHead(UserIndex, "No tenes esa cantidad.", Npclist(.flags.TargetNPC).Char.CharIndex, 255, 255, 255)

        End If

    End With

End Sub

''
' Handles the "Denounce" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDenounce(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 14/11/2010
    '14/11/2010: ZaMa - Now denounces can be desactivated.
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim Text As String

        Dim msg  As String
        
        Text = Buffer.ReadASCIIString()
        
        If .flags.Silenciado = 0 Then
            'Analize chat...
            Call Statistics.ParseChat(Text)
            
            msg = LCase$(.Name) & " DENUNCIA: " & Text
            
            Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(msg, FontTypeNames.FONTTYPE_GUILDMSG), True)
            
            Call Denuncias.Push(msg, False)
            
            Call WriteConsoleMsg(UserIndex, "Denuncia enviada, espere..", FontTypeNames.FONTTYPE_INFO)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "GuildFundate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildFundate(ByVal UserIndex As Integer)

    '***************************************************
    'Author: ZaMa
    'Last Modification: 14/12/2009
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 1 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        Call .incomingData.ReadByte
        
        If EsGm(UserIndex) Or EsRolesMaster(UserList(UserIndex).Name) Then
            Call WriteConsoleMsg(UserIndex, "Los GM's no pueden fundar clanes.", FontTypeNames.FONTTYPE_INFOBOLD)
            Exit Sub

        End If
        
        If HasFound(.Name) Then
            Call WriteConsoleMsg(UserIndex, "Ya has fundado un clan, no puedes fundar otro!", FontTypeNames.FONTTYPE_INFOBOLD)
            Exit Sub

        End If
        
        Call WriteShowGuildAlign(UserIndex)

    End With

End Sub
    
''
' Handles the "GuildFundation" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildFundation(ByVal UserIndex As Integer)

    '***************************************************
    'Author: ZaMa
    'Last Modification: 14/12/2009
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim clanType As eClanType

        Dim Error    As String
        
        clanType = .incomingData.ReadByte()
        
        If HasFound(.Name) Then
            Call WriteConsoleMsg(UserIndex, "Ya has fundado un clan, no puedes fundar otro!", FontTypeNames.FONTTYPE_INFOBOLD)
            Call LogCheating("El usuario " & .Name & " ha intentado fundar un clan ya habiendo fundado otro desde la IP " & .IP)
            Exit Sub

        End If
        
        Select Case UCase$(Trim(clanType))

            Case eClanType.ct_RoyalArmy
                .FundandoGuildAlineacion = ALINEACION_ARMADA

            Case eClanType.ct_Evil
                .FundandoGuildAlineacion = ALINEACION_LEGION

            Case eClanType.ct_Neutral
                .FundandoGuildAlineacion = ALINEACION_NEUTRO

            Case eClanType.ct_Legal
                .FundandoGuildAlineacion = ALINEACION_CIUDA

            Case eClanType.ct_Criminal
                .FundandoGuildAlineacion = ALINEACION_CRIMINAL

            Case Else
                Call WriteConsoleMsg(UserIndex, "Alineacion invalida.", FontTypeNames.FONTTYPE_GUILD)
                Exit Sub

        End Select
        
        If modGuilds.PuedeFundarUnClan(UserIndex, .FundandoGuildAlineacion, Error) Then
            Call WriteShowGuildFundationForm(UserIndex)
        Else
            .FundandoGuildAlineacion = 0
            Call WriteConsoleMsg(UserIndex, Error, FontTypeNames.FONTTYPE_GUILD)

        End If

    End With

End Sub

''
' Handles the "PartyKick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePartyKick(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/05/09
    'Last Modification by: Marco Vanotti (Marco)
    '- 05/05/09: Now it uses "UserPuedeEjecutarComandos" to check if the user can use party commands
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim username As String

        Dim tUser    As Integer
        
        username = Buffer.ReadASCIIString()
        
        If UserPuedeEjecutarComandos(UserIndex) Then
            tUser = NameIndex(username)
            
            If tUser > 0 Then
                Call mdParty.ExpulsarDeParty(UserIndex, tUser)
            Else

                If InStr(username, "+") Then
                    username = Replace(username, "+", " ")

                End If
                
                Call WriteConsoleMsg(UserIndex, LCase(username) & " no pertenece a tu party.", FontTypeNames.FONTTYPE_INFO)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "PartySetLeader" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePartySetLeader(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/05/09
    'Last Modification by: Marco Vanotti (MarKoxX)
    '- 05/05/09: Now it uses "UserPuedeEjecutarComandos" to check if the user can use party commands
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'On Error GoTo ErrHandler
    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim username As String

        Dim tUser    As Integer

        Dim Rank     As Integer

        Rank = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero
        
        username = Buffer.ReadASCIIString()

        If UserPuedeEjecutarComandos(UserIndex) Then
            tUser = NameIndex(username)

            If tUser > 0 Then

                'Don't allow users to spoof online GMs
                If (UserDarPrivilegioLevel(username) And Rank) <= (.flags.Privilegios And Rank) Then
                    Call mdParty.TransformarEnLider(UserIndex, tUser)
                Else
                    Call WriteConsoleMsg(UserIndex, LCase(UserList(tUser).Name) & " no pertenece a tu party.", FontTypeNames.FONTTYPE_INFO)

                End If
                
            Else

                If InStr(username, "+") Then
                    username = Replace(username, "+", " ")

                End If

                Call WriteConsoleMsg(UserIndex, LCase(username) & " no pertenece a tu party.", FontTypeNames.FONTTYPE_INFO)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "PartyAcceptMember" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePartyAcceptMember(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/05/09
    'Last Modification by: Marco Vanotti (Marco)
    '- 05/05/09: Now it uses "UserPuedeEjecutarComandos" to check if the user can use party commands
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim username  As String

        Dim tUser     As Integer

        Dim Rank      As Integer

        Dim bUserVivo As Boolean
        
        Rank = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero
        
        username = Buffer.ReadASCIIString()

        If UserList(UserIndex).flags.Muerto Then
            Call WriteConsoleMsg(UserIndex, "Estas muerto!!", FontTypeNames.FONTTYPE_PARTY)
        Else
            bUserVivo = True

        End If
        
        If mdParty.UserPuedeEjecutarComandos(UserIndex) And bUserVivo Then
            tUser = NameIndex(username)

            If tUser > 0 Then

                'Validate administrative ranks - don't allow users to spoof online GMs
                If (UserList(tUser).flags.Privilegios And Rank) <= (.flags.Privilegios And Rank) Then
                    Call mdParty.AprobarIngresoAParty(UserIndex, tUser)
                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes incorporar a tu party a personajes de mayor jerarquia.", FontTypeNames.FONTTYPE_INFO)

                End If

            Else

                If InStr(username, "+") Then
                    username = Replace(username, "+", " ")

                End If
                
                'Don't allow users to spoof online GMs
                If (UserDarPrivilegioLevel(username) And Rank) <= (.flags.Privilegios And Rank) Then
                    Call WriteConsoleMsg(UserIndex, LCase(username) & " no ha solicitado ingresar a tu party.", FontTypeNames.FONTTYPE_PARTY)
                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes incorporar a tu party a personajes de mayor jerarquia.", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "GuildMemberList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildMemberList(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim Guild       As String

        Dim memberCount As Integer

        Dim i           As Long

        Dim username    As String
        
        Guild = Buffer.ReadASCIIString()
        
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then
            If (InStrB(Guild, "\") <> 0) Then
                Guild = Replace(Guild, "\", "")

            End If

            If (InStrB(Guild, "/") <> 0) Then
                Guild = Replace(Guild, "/", "")

            End If
            
            If Not FileExist(App.Path & "\guilds\" & Guild & "-members.mem") Then
                Call WriteConsoleMsg(UserIndex, "No existe el clan: " & Guild, FontTypeNames.FONTTYPE_INFO)
            Else
                memberCount = val(GetVar(App.Path & "\Guilds\" & Guild & "-Members" & ".mem", "INIT", "NroMembers"))
                
                For i = 1 To memberCount
                    username = GetVar(App.Path & "\Guilds\" & Guild & "-Members" & ".mem", "Members", "Member" & i)
                    
                    Call WriteConsoleMsg(UserIndex, username & "<" & Guild & ">", FontTypeNames.FONTTYPE_INFO)
                Next i

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "GMMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGMMessage(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 01/08/07
    'Last Modification by: (liquid)
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim Message As String
        
        Message = Buffer.ReadASCIIString()
        
        If Not .flags.Privilegios And PlayerType.User Then
            Call LogGM(.Name, "Mensaje a Gms:" & Message)
        
            If LenB(Message) <> 0 Then
                'Analize chat...
                Call Statistics.ParseChat(Message)
            
                Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.Name & "> " & Message, FontTypeNames.FONTTYPE_GMMSG))

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "ShowName" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleShowName(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            .showName = Not .showName 'Show / Hide the name
            
            Call RefreshCharStatus(UserIndex)

        End If

    End With

End Sub

''
' Handles the "OnlineRoyalArmy" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleOnlineRoyalArmy(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 28/05/2010
    '28/05/2010: ZaMa - Ahora solo dioses pueden ver otros dioses online.
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        .incomingData.ReadByte
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
    
        Dim i    As Long

        Dim list As String

        Dim priv As PlayerType

        priv = PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios
        
        ' Solo dioses pueden ver otros dioses online
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
            priv = priv Or PlayerType.Dios Or PlayerType.Admin

        End If
     
        For i = 1 To LastUser

            If UserList(i).ConnID <> -1 Then
                If UserList(i).Faccion.ArmadaReal = 1 Then
                    If UserList(i).flags.Privilegios And priv Then
                        list = list & UserList(i).Name & ", "

                    End If

                End If

            End If

        Next i

    End With
    
    If Len(list) > 0 Then
        Call WriteConsoleMsg(UserIndex, "Reales conectados: " & Left$(list, Len(list) - 2), FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(UserIndex, "No hay reales conectados.", FontTypeNames.FONTTYPE_INFO)

    End If

End Sub

''
' Handles the "OnlineChaosLegion" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleOnlineChaosLegion(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 28/05/2010
    '28/05/2010: ZaMa - Ahora solo dioses pueden ver otros dioses online.
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        .incomingData.ReadByte
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
    
        Dim i    As Long

        Dim list As String

        Dim priv As PlayerType

        priv = PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios
        
        ' Solo dioses pueden ver otros dioses online
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
            priv = priv Or PlayerType.Dios Or PlayerType.Admin

        End If
     
        For i = 1 To LastUser

            If UserList(i).ConnID <> -1 Then
                If UserList(i).Faccion.FuerzasCaos = 1 Then
                    If UserList(i).flags.Privilegios And priv Then
                        list = list & UserList(i).Name & ", "

                    End If

                End If

            End If

        Next i

    End With

    If Len(list) > 0 Then
        Call WriteConsoleMsg(UserIndex, "Caos conectados: " & Left$(list, Len(list) - 2), FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(UserIndex, "No hay Caos conectados.", FontTypeNames.FONTTYPE_INFO)

    End If

End Sub

''
' Handles the "GoNearby" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGoNearby(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 01/10/07
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim username As String
        
        username = Buffer.ReadASCIIString()
        
        Dim tIndex As Integer

        Dim X      As Long

        Dim Y      As Long

        Dim i      As Long

        Dim Found  As Boolean
        
        tIndex = NameIndex(username)
        
        'Check the user has enough powers
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero) Then

            'Si es dios o Admins no podemos salvo que nosotros tambien lo seamos
            If Not (EsDios(username) Or EsAdmin(username)) Or (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) Then
                If tIndex <= 0 Then 'existe el usuario destino?
                    Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
                Else

                    For i = 2 To 5 'esto for sirve ir cambiando la distancia destino
                        For X = UserList(tIndex).Pos.X - i To UserList(tIndex).Pos.X + i
                            For Y = UserList(tIndex).Pos.Y - i To UserList(tIndex).Pos.Y + i

                                If MapData(UserList(tIndex).Pos.Map, X, Y).UserIndex = 0 Then
                                    If LegalPos(UserList(tIndex).Pos.Map, X, Y, True, True) Then
                                        Call WarpUserChar(UserIndex, UserList(tIndex).Pos.Map, X, Y, True)
                                        Call LogGM(.Name, "/IRCERCA " & username & " Mapa:" & UserList(tIndex).Pos.Map & " X:" & UserList(tIndex).Pos.X & " Y:" & UserList(tIndex).Pos.Y)
                                        Found = True
                                        Exit For

                                    End If

                                End If

                            Next Y
                            
                            If Found Then Exit For  ' Feo, pero hay que abortar 3 fors sin usar GoTo
                        Next X
                        
                        If Found Then Exit For  ' Feo, pero hay que abortar 3 fors sin usar GoTo
                    Next i
                    
                    'No space found??
                    If Not Found Then
                        Call WriteConsoleMsg(UserIndex, "Todos los lugares estan ocupados.", FontTypeNames.FONTTYPE_INFO)

                    End If

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "Comment" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleComment(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim comment As String

        comment = Buffer.ReadASCIIString()
        
        If Not .flags.Privilegios And PlayerType.User Then
            Call LogGM(.Name, "Comentario: " & comment)
            Call WriteConsoleMsg(UserIndex, "Comentario salvado...", FontTypeNames.FONTTYPE_INFO)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "ServerTime" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleServerTime(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 01/08/07
    'Last Modification by: (liquid)
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
    
        If .flags.Privilegios And PlayerType.User Then Exit Sub
    
        Call LogGM(.Name, "Hora.")

    End With
    
    Call modSendData.SendData(SendTarget.Toall, 0, PrepareMessageConsoleMsg("Hora: " & time & " " & Date, FontTypeNames.FONTTYPE_INFO))

End Sub

''
' Handles the "Where" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWhere(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 18/11/2010
    '07/06/2010: ZaMa - Ahora no se puede usar para saber si hay dioses/admins online.
    '18/11/2010: ZaMa - Obtengo los privs del charfile antes de mostrar la posicion de un usuario offline.
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim username As String

        Dim tUser    As Integer

        Dim miPos    As String
        
        username = Buffer.ReadASCIIString()
        
        If Not .flags.Privilegios And PlayerType.User Then
            
            tUser = NameIndex(username)

            If tUser <= 0 Then
                
                If PersonajeExiste(username) Then
                
                    Dim CharPrivs As PlayerType

                    CharPrivs = GetCharPrivs(username)
                    
                    If (CharPrivs And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios)) <> 0 Or ((CharPrivs And (PlayerType.Dios Or PlayerType.Admin) <> 0) And (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0) Then
                        miPos = GetUserPos(username)
                        Call WriteConsoleMsg(UserIndex, "Ubicacion  " & username & " (Offline): " & miPos & ".", FontTypeNames.FONTTYPE_INFO)

                    End If

                Else

                    If Not (EsDios(username) Or EsAdmin(username)) Then
                        Call WriteConsoleMsg(UserIndex, "Usuario inexistente.", FontTypeNames.FONTTYPE_INFO)
                    ElseIf .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
                        Call WriteConsoleMsg(UserIndex, "Usuario inexistente.", FontTypeNames.FONTTYPE_INFO)

                    End If

                End If

            Else

                If (UserList(tUser).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios)) <> 0 Or ((UserList(tUser).flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) <> 0) And (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0) Then
                    Call WriteConsoleMsg(UserIndex, "Ubicacion  " & username & ": " & UserList(tUser).Pos.Map & ", " & UserList(tUser).Pos.X & ", " & UserList(tUser).Pos.Y & ".", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If
        
        Call LogGM(.Name, "/Donde " & username)
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "CreaturesInMap" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCreaturesInMap(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 30/07/06
    'Pablo (ToxicWaste): modificaciones generales para simplificar la visualizacion.
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Map As Integer

        Dim i, j As Long

        Dim NPCcount1, NPCcount2 As Integer

        Dim NPCcant1() As Integer

        Dim NPCcant2() As Integer

        Dim List1()    As String

        Dim List2()    As String
        
        Map = .incomingData.ReadInteger()
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
        
        If MapaValido(Map) Then

            For i = 1 To LastNPC

                'VB isn't lazzy, so we put more restrictive condition first to speed up the process
                If NPCZonaId(i) = UserZonaId(UserIndex) Then

                    'esta vivo?
                    If Npclist(i).flags.NPCActive And Npclist(i).Hostile = 1 And Npclist(i).Stats.Alineacion = 2 Then
                        If NPCcount1 = 0 Then
                            ReDim List1(0) As String
                            ReDim NPCcant1(0) As Integer
                            NPCcount1 = 1
                            List1(0) = Npclist(i).Name & ": (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
                            NPCcant1(0) = 1
                        Else

                            For j = 0 To NPCcount1 - 1

                                If Left$(List1(j), Len(Npclist(i).Name)) = Npclist(i).Name Then
                                    List1(j) = List1(j) & ", (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
                                    NPCcant1(j) = NPCcant1(j) + 1
                                    Exit For

                                End If

                            Next j

                            If j = NPCcount1 Then
                                ReDim Preserve List1(0 To NPCcount1) As String
                                ReDim Preserve NPCcant1(0 To NPCcount1) As Integer
                                NPCcount1 = NPCcount1 + 1
                                List1(j) = Npclist(i).Name & ": (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
                                NPCcant1(j) = 1

                            End If

                        End If

                    Else

                        If NPCcount2 = 0 Then
                            ReDim List2(0) As String
                            ReDim NPCcant2(0) As Integer
                            NPCcount2 = 1
                            List2(0) = Npclist(i).Name & ": (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
                            NPCcant2(0) = 1
                        Else

                            For j = 0 To NPCcount2 - 1

                                If Left$(List2(j), Len(Npclist(i).Name)) = Npclist(i).Name Then
                                    List2(j) = List2(j) & ", (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
                                    NPCcant2(j) = NPCcant2(j) + 1
                                    Exit For

                                End If

                            Next j

                            If j = NPCcount2 Then
                                ReDim Preserve List2(0 To NPCcount2) As String
                                ReDim Preserve NPCcant2(0 To NPCcount2) As Integer
                                NPCcount2 = NPCcount2 + 1
                                List2(j) = Npclist(i).Name & ": (" & Npclist(i).Pos.X & "," & Npclist(i).Pos.Y & ")"
                                NPCcant2(j) = 1

                            End If

                        End If

                    End If

                End If

            Next i
            
            Call WriteConsoleMsg(UserIndex, "Npcs Hostiles en zona: ", FontTypeNames.FONTTYPE_WARNING)

            If NPCcount1 = 0 Then
                Call WriteConsoleMsg(UserIndex, "No hay NPCS Hostiles.", FontTypeNames.FONTTYPE_INFO)
            Else

                For j = 0 To NPCcount1 - 1
                    Call WriteConsoleMsg(UserIndex, NPCcant1(j) & " " & List1(j), FontTypeNames.FONTTYPE_INFO)
                Next j

            End If

            Call WriteConsoleMsg(UserIndex, "Otros Npcs en la Zona: ", FontTypeNames.FONTTYPE_WARNING)

            If NPCcount2 = 0 Then
                Call WriteConsoleMsg(UserIndex, "No hay mas NPCS.", FontTypeNames.FONTTYPE_INFO)
            Else

                For j = 0 To NPCcount2 - 1
                    Call WriteConsoleMsg(UserIndex, NPCcant2(j) & " " & List2(j), FontTypeNames.FONTTYPE_INFO)
                Next j

            End If

            Call LogGM(.Name, "Numero enemigos en zona " & Map)

        End If

    End With

End Sub

''
' Handles the "WarpMeToTarget" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWarpMeToTarget(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 26/03/09
    '26/03/06: ZaMa - Chequeo que no se teletransporte donde haya un char o npc
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim X As Integer

        Dim Y As Integer
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
        
        X = .flags.TargetX
        Y = .flags.TargetY
        
        Call FindLegalPos(UserIndex, .flags.TargetMap, X, Y)
        Call WarpUserChar(UserIndex, .flags.TargetMap, X, Y, True)
        Call LogGM(.Name, "/TELEPLOC a x:" & .flags.TargetX & " Y:" & .flags.TargetY & " Map:" & .Pos.Map)

    End With

End Sub

''
' Handles the "WarpChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWarpChar(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 11/08/2019
    '26/03/2009: ZaMa - Chequeo que no se teletransporte a un tile donde haya un char o npc.
    '11/08/2019: Jopi - No registramos en los logs si te teletransportas a vos mismo.
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 9 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim username  As String

        Dim Map       As Integer

        Dim X         As Integer

        Dim Y         As Integer
        
        Dim Cuadrante As Boolean

        Dim tUser     As Integer
        
        username = Buffer.ReadASCIIString()
        Map = Buffer.ReadInteger()
        X = Buffer.ReadInteger()
        Y = Buffer.ReadInteger()
        Cuadrante = Buffer.ReadBoolean()
        
        If Not .flags.Privilegios And PlayerType.User Then
            If LenB(username) <> 0 Then
                If UCase$(username) <> "YO" Then
                    If Not .flags.Privilegios And PlayerType.Consejero Then
                        tUser = NameIndex(username)

                    End If

                Else
                    tUser = UserIndex

                End If
            
                If tUser <= 0 Then
                    If Not (EsDios(username) Or EsAdmin(username)) Then
                        Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
                    Else
                        Call WriteConsoleMsg(UserIndex, "No puedes transportar dioses o admins.", FontTypeNames.FONTTYPE_INFO)

                    End If
                    
                ElseIf Not ((UserList(tUser).flags.Privilegios And PlayerType.Dios) <> 0 Or (UserList(tUser).flags.Privilegios And PlayerType.Admin) <> 0) Or tUser = UserIndex Then
                            
                    'Si es un TP por cuadrante calculamos las coord y las reemplazamos
                    If Cuadrante Then
                    
                        Dim tmpX As Integer
                        Dim tmpY As Integer
                        
                        Call ObtenerCoordenadasDesdeCuadrante(Map, X, Y, tmpX, tmpY)
                        
                        Map = UserList(UserIndex).Pos.Map
                        X = tmpX
                        Y = tmpY
                    
                    End If
                    
                    If MapaValido(Map) Then
                        If InMapBounds(Map, X, Y) Then
                            Call FindLegalPos(tUser, Map, X, Y)
                            Call WarpUserChar(tUser, Map, X, Y, True, True)
                            
                            ' Agrego esto para no llenar consola de mensajes al hacer SHIFT + CLICK DERECHO
                            If UserIndex <> tUser Then
                                Call WriteConsoleMsg(UserIndex, UserList(tUser).Name & " transportado.", FontTypeNames.FONTTYPE_INFO)
                                Call LogGM(.Name, "Transporto a " & UserList(tUser).Name & " hacia " & "Mapa" & Map & " X:" & X & " Y:" & Y)
    
                            End If
                        End If
                    End If

                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes transportar dioses o admins.", FontTypeNames.FONTTYPE_INFO)
            
                End If
            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "Silence" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSilence(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim username As String

        Dim tUser    As Integer
        
        username = Buffer.ReadASCIIString()
        
        If Not .flags.Privilegios And PlayerType.User Then
            tUser = NameIndex(username)
        
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
            Else

                If UserList(tUser).flags.Silenciado = 0 Then
                    UserList(tUser).flags.Silenciado = 1
                    Call WriteConsoleMsg(UserIndex, "Usuario silenciado.", FontTypeNames.FONTTYPE_INFO)
                    Call WriteShowMessageBox(tUser, "Estimado usuario, ud. ha sido silenciado por los administradores. Sus denuncias seran ignoradas por el servidor de aqui en mas. Utilice /GM para contactar un administrador.")
                    Call LogGM(.Name, "/silenciar " & UserList(tUser).Name)
                Else
                    UserList(tUser).flags.Silenciado = 0
                    Call WriteConsoleMsg(UserIndex, "Usuario des silenciado.", FontTypeNames.FONTTYPE_INFO)
                    Call LogGM(.Name, "/DESsilenciar " & UserList(tUser).Name)

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "SOSShowList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSOSShowList(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
        Call WriteShowSOSForm(UserIndex)

    End With

End Sub

''
' Handles the "RequestPartyForm" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandlePartyForm(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Budi
    'Last Modification: 11/26/09
    '
    '***************************************************
    
    Dim LiderInvita As Boolean
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        LiderInvita = .incomingData.ReadBoolean

        If LiderInvita Then
            Call WritePeticionInvitarParty(UserIndex)

        ElseIf .PartyIndex > 0 Then
            Call WriteShowPartyForm(UserIndex)
            
        Else
            Call WritePeticionInvitarParty(UserIndex)

        End If

    End With

End Sub

''
' Handles the "SOSRemove" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSOSRemove(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim username As String

        username = Buffer.ReadASCIIString()
        
        If Not .flags.Privilegios And PlayerType.User Then Call Ayuda.Quitar(username)
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "GoToChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGoToChar(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 26/03/2009
    '26/03/2009: ZaMa -  Chequeo que no se teletransporte a un tile donde haya un char o npc.
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim username As String

        Dim tUser    As Integer

        Dim X        As Integer

        Dim Y        As Integer
        
        username = Buffer.ReadASCIIString()
        tUser = NameIndex(username)
        
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.SemiDios Or PlayerType.Consejero) Then

            'Si es dios o Admins no podemos salvo que nosotros tambien lo seamos
            If Not (EsDios(username) Or EsAdmin(username)) Or (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0 Then
                If tUser <= 0 Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
                Else
                    X = UserList(tUser).Pos.X
                    Y = UserList(tUser).Pos.Y + 1
                    Call FindLegalPos(UserIndex, UserList(tUser).Pos.Map, X, Y)
                    
                    Call WarpUserChar(UserIndex, UserList(tUser).Pos.Map, X, Y, True)
                    
                    If .flags.AdminInvisible = 0 Then
                        Call WriteConsoleMsg(tUser, .Name & " se ha trasportado hacia donde te encuentras.", FontTypeNames.FONTTYPE_INFO)

                    End If
                    
                    Call LogGM(.Name, "/IRA " & username & " Mapa:" & UserList(tUser).Pos.Map & " X:" & UserList(tUser).Pos.X & " Y:" & UserList(tUser).Pos.Y)

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "Invisible" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleInvisible(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
        
        Call DoAdminInvisible(UserIndex)
        Call LogGM(.Name, "/INVISIBLE")

    End With

End Sub

''
' Handles the "GMPanel" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGMPanel(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Dim ID As Byte
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        ID = .incomingData.ReadByte
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
        
        Call WriteShowGMPanelForm(UserIndex, ID)

    End With

End Sub

''
' Handles the "GMPanel" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestUserList(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 01/09/07
    'Last modified by: Lucas Tavolaro Ortiz (Tavo)
    'I haven`t found a solution to split, so i make an array of names
    '***************************************************
    Dim i       As Long

    Dim names() As String

    Dim Count   As Long
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.RoleMaster) Then Exit Sub
        
        ReDim names(1 To LastUser) As String
        Count = 1
        
        For i = 1 To LastUser

            If (LenB(UserList(i).Name) <> 0) Then
                If UserList(i).flags.Privilegios And PlayerType.User Then
                    names(Count) = UserList(i).Name
                    Count = Count + 1

                End If

            End If

        Next i
        
        If Count > 1 Then Call WriteUserNameList(UserIndex, names(), Count - 1)

    End With

End Sub

''
' Handles the "Working" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWorking(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 07/10/2010
    '***************************************************
    Dim i     As Long

    Dim Users As String
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.RoleMaster) Then Exit Sub
        
        For i = 1 To LastUser

            If UserList(i).flags.UserLogged And UserList(i).Counters.Trabajando > 0 Then
                Users = Users & ", " & UserList(i).Name

            End If

        Next i
        
        If LenB(Users) <> 0 Then
            Users = Right$(Users, Len(Users) - 2)
            Call WriteConsoleMsg(UserIndex, "Usuarios trabajando: " & Users, FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(UserIndex, "No hay usuarios trabajando.", FontTypeNames.FONTTYPE_INFO)

        End If

    End With

End Sub

''
' Handles the "Hiding" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleHiding(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    Dim i     As Long

    Dim Users As String
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.RoleMaster) Then Exit Sub
        
        For i = 1 To LastUser

            If (LenB(UserList(i).Name) <> 0) And UserList(i).Counters.Ocultando > 0 Then
                Users = Users & UserList(i).Name & ", "

            End If

        Next i
        
        If LenB(Users) <> 0 Then
            Users = Left$(Users, Len(Users) - 2)
            Call WriteConsoleMsg(UserIndex, "Usuarios ocultandose: " & Users, FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(UserIndex, "No hay usuarios ocultandose.", FontTypeNames.FONTTYPE_INFO)

        End If

    End With

End Sub

''
' Handles the "Jail" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleJail(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 07/06/2010
    '07/06/2010: ZaMa - Ahora no se puede usar para saber si hay dioses/admins online.
    'Last Modification: 04/04/2020
    '4/4/2020: FrankoH298 - Ahora calcula bien el tiempo de carcel
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 6 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim username As String

        Dim Reason   As String

        Dim jailTime As Byte

        Dim Count    As Byte

        Dim tUser    As Integer
        
        username = Buffer.ReadASCIIString()
        Reason = Buffer.ReadASCIIString()
        jailTime = Buffer.ReadByte()
        
        If InStr(1, username, "+") Then
            username = Replace(username, "+", " ")

        End If
        
        '/carcel nick@motivo@<tiempo>
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (Not .flags.Privilegios And PlayerType.User) <> 0 Then
            If LenB(username) = 0 Or LenB(Reason) = 0 Then
                Call WriteConsoleMsg(UserIndex, "Utilice /carcel nick@motivo@tiempo", FontTypeNames.FONTTYPE_INFO)
            Else
                tUser = NameIndex(username)
                
                If tUser <= 0 Then
                    If (EsDios(username) Or EsAdmin(username)) Then
                        Call WriteConsoleMsg(UserIndex, "No puedes encarcelar a administradores.", FontTypeNames.FONTTYPE_INFO)
                    Else
                        Call WriteConsoleMsg(UserIndex, "El usuario no esta online.", FontTypeNames.FONTTYPE_INFO)

                    End If

                Else

                    If Not UserList(tUser).flags.Privilegios And PlayerType.User Then
                        Call WriteConsoleMsg(UserIndex, "No puedes encarcelar a administradores.", FontTypeNames.FONTTYPE_INFO)
                    ElseIf jailTime > (60) Then
                        Call WriteConsoleMsg(UserIndex, "No puedes encarcelar por mas de 60 minutos.", FontTypeNames.FONTTYPE_INFO)
                    Else

                        If (InStrB(username, "\") <> 0) Then
                            username = Replace(username, "\", "")

                        End If

                        If (InStrB(username, "/") <> 0) Then
                            username = Replace(username, "/", "")

                        End If
                        
                        If PersonajeExiste(username) Then
                            Count = GetUserAmountOfPunishments(username)
                            Call SaveUserPunishment(username, Count + 1, LCase$(.Name) & ": CARCEL " & jailTime & "m, MOTIVO: " & LCase$(Reason) & " " & Date & " " & time)

                        End If
                        
                        Call Encarcelar(tUser, jailTime, .Name)
                        Call LogGM(.Name, " encarcelo a " & username)

                    End If

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "KillNPC" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleKillNPC(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 04/22/08 (NicoNZ)
    '
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
        
        Dim tNPC   As Integer

        Dim auxNPC As NPC
        
        'Los consejeros no pueden RMATAr a nada en el mapa pretoriano
        If .flags.Privilegios And PlayerType.Consejero Then
            If .Pos.Map = MAPA_PRETORIANO Then
                Call WriteConsoleMsg(UserIndex, "Los consejeros no pueden usar este comando en el mapa pretoriano.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If

        End If
        
        tNPC = .flags.TargetNPC
        
        If tNPC > 0 Then
            Call WriteConsoleMsg(UserIndex, "RMatas (con posible respawn) a: " & Npclist(tNPC).Name, FontTypeNames.FONTTYPE_INFO)
            
            auxNPC = Npclist(tNPC)
            Call QuitarNPC(tNPC)
            Call ReSpawnNpc(auxNPC)
            
            .flags.TargetNPC = 0
        Else
            Call WriteConsoleMsg(UserIndex, "Antes debes hacer click sobre el NPC.", FontTypeNames.FONTTYPE_INFO)

        End If

    End With

End Sub

''
' Handles the "WarnUser" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleWarnUser(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/26/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim username As String

        Dim Reason   As String

        Dim Privs    As PlayerType

        Dim Count    As Byte
        
        username = Buffer.ReadASCIIString()
        Reason = Buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (Not .flags.Privilegios And PlayerType.User) <> 0 Then
            If LenB(username) = 0 Or LenB(Reason) = 0 Then
                Call WriteConsoleMsg(UserIndex, "Utilice /advertencia nick@motivo", FontTypeNames.FONTTYPE_INFO)
            Else
                Privs = UserDarPrivilegioLevel(username)
                
                If Not Privs And PlayerType.User Then
                    Call WriteConsoleMsg(UserIndex, "No puedes advertir a administradores.", FontTypeNames.FONTTYPE_INFO)
                Else

                    If (InStrB(username, "\") <> 0) Then
                        username = Replace(username, "\", "")

                    End If

                    If (InStrB(username, "/") <> 0) Then
                        username = Replace(username, "/", "")

                    End If
                    
                    If PersonajeExiste(username) Then
                        Count = GetUserAmountOfPunishments(username)
                        Call SaveUserPunishment(username, Count + 1, LCase$(.Name) & ": ADVERTENCIA por: " & LCase$(Reason) & " " & Date & " " & time)

                        Call WriteConsoleMsg(UserIndex, "Has advertido a " & UCase$(username) & ".", FontTypeNames.FONTTYPE_INFO)
                        Call LogGM(.Name, " advirtio a " & username)

                    End If

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "EditChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleEditChar(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 11/05/2019
    '02/03/2009: ZaMa - Cuando editas nivel, chequea si el pj puede permanecer en clan faccionario
    '11/06/2009: ZaMa - Todos los comandos se pueden usar aunque el pj este offline
    '18/09/2010: ZaMa - Ahora se puede editar la vida del propio pj (cualquier rm o dios).
    '11/05/2019: Jopi - No registramos en los logs si te editas a vos mismo.
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 8 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim username      As String

        Dim tUser         As Integer

        Dim Opcion        As Byte

        Dim Arg1          As String

        Dim Arg2          As String

        Dim valido        As Boolean

        Dim LoopC         As Byte

        Dim CommandString As String

        Dim n             As Byte

        Dim Var           As Long
        
        username = Replace(Buffer.ReadASCIIString(), "+", " ")
        
        If UCase$(username) = "YO" Then
            tUser = UserIndex
        Else
            tUser = NameIndex(username)

        End If
        
        Opcion = Buffer.ReadByte()
        Arg1 = Buffer.ReadASCIIString()
        Arg2 = Buffer.ReadASCIIString()
        
        If .flags.Privilegios And PlayerType.RoleMaster Then

            Select Case .flags.Privilegios And (PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero)

                Case PlayerType.Consejero
                    ' Los RMs consejeros solo se pueden editar su head, body, level y vida
                    valido = tUser = UserIndex And (Opcion = eEditOptions.eo_Body Or Opcion = eEditOptions.eo_Head Or Opcion = eEditOptions.eo_Level Or Opcion = eEditOptions.eo_Vida)
                
                Case PlayerType.SemiDios
                    ' Los RMs solo se pueden editar su level o vida y el head y body de cualquiera
                    valido = ((Opcion = eEditOptions.eo_Level Or Opcion = eEditOptions.eo_Vida) And tUser = UserIndex) Or Opcion = eEditOptions.eo_Body Or Opcion = eEditOptions.eo_Head
                    
                Case PlayerType.Dios
                    ' Los DRMs pueden aplicar los siguientes comandos sobre cualquiera
                    ' pero si quiere modificar el level o vida solo lo puede hacer sobre si mismo
                    valido = ((Opcion = eEditOptions.eo_Level Or Opcion = eEditOptions.eo_Vida) And tUser = UserIndex) Or Opcion = eEditOptions.eo_Body Or Opcion = eEditOptions.eo_Head Or Opcion = eEditOptions.eo_CiticensKilled Or Opcion = eEditOptions.eo_CriminalsKilled Or Opcion = eEditOptions.eo_Class Or Opcion = eEditOptions.eo_Skills Or Opcion = eEditOptions.eo_addGold

            End Select
        
            'Si no es RM debe ser dios para poder usar este comando
        ElseIf .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then
            
            If Opcion = eEditOptions.eo_Vida Then
                '  Por ahora dejo para que los dioses no puedan editar la vida de otros
                valido = (tUser = UserIndex)
            Else
                valido = True

            End If
            
        ElseIf .flags.PrivEspecial Then
            valido = (Opcion = eEditOptions.eo_CiticensKilled) Or (Opcion = eEditOptions.eo_CriminalsKilled)
            
        End If

        'CHOTS | The user is not online and we are working with Database
        If tUser <= 0 Then
            valido = False
            Call WriteConsoleMsg(UserIndex, "El usuario esta offline.", FontTypeNames.FONTTYPE_INFO)

            '@TODO call a method to edit the user using the database
        End If

        If valido Then
            'For making the Log
            CommandString = "/MOD "
                
            Select Case Opcion

                Case eEditOptions.eo_Gold

                    If val(Arg1) <= MAX_ORO_EDIT Then
                        If tUser <= 0 Then ' Esta offline?
                            Call WriteConsoleMsg(UserIndex, "El usuario esta offline o no existe.", FontTypeNames.FONTTYPE_INFO)
                            Call LogGM(.Name, "Intento editar un usuario inexistente u offline.")
                        Else ' Online
                            UserList(tUser).Stats.Gld = val(Arg1)
                            Call WriteUpdateGold(tUser)

                        End If

                    Else
                            Call WriteConsoleMsg(UserIndex, "No esta permitido utilizar valores mayores a " & MAX_ORO_EDIT & ". Su comando ha quedado en los logs del juego.", FontTypeNames.FONTTYPE_INFO)

                    End If
                    
                    ' Log it
                    CommandString = CommandString & "ORO "
                
                Case eEditOptions.eo_Experience

                        If val(Arg1) <= MAX_EXP_EDIT Then
                        
                            If tUser <= 0 Then ' Offline
                                Call WriteConsoleMsg(UserIndex, "El usuario esta offline o no existe.", FontTypeNames.FONTTYPE_INFO)
                                Call LogGM(.Name, "Intento editar un usuario inexistente u offline.")
                            Else ' Online
                                UserList(tUser).Stats.Exp = UserList(tUser).Stats.Exp + val(Arg1)
                                Call CheckUserLevel(tUser)
                                Call WriteUpdateExp(tUser)
    
                            End If
                        Else
                            Call WriteConsoleMsg(UserIndex, "No esta permitido utilizar valores mayores a " & MAX_EXP_EDIT & ". Su comando ha quedado en los logs del juego.", FontTypeNames.FONTTYPE_INFO)

                        End If
                        
                        ' Log it
                        CommandString = CommandString & "EXP "
                        
                    Case eEditOptions.eo_ExperiencePVP

                        If val(Arg1) <= MAX_EXP_EDIT Then
                        
                            If tUser <= 0 Then ' Offline
                                Call WriteConsoleMsg(UserIndex, "El usuario esta offline o no existe.", FontTypeNames.FONTTYPE_INFO)
                                Call LogGM(.Name, "Intento editar un usuario inexistente u offline.")
                            Else ' Online
                                UserList(tUser).Stats.ExpPVP = UserList(tUser).Stats.ExpPVP + val(Arg1)
                                Call CheckUserLevelPVP(tUser)
    
                            End If
                        Else
                            Call WriteConsoleMsg(UserIndex, "No esta permitido utilizar valores mayores a " & MAX_EXP_EDIT & ". Su comando ha quedado en los logs del juego.", FontTypeNames.FONTTYPE_INFO)

                        End If
                        
                        ' Log it
                        CommandString = CommandString & "EXPPVP "
                    
                Case eEditOptions.eo_Body

                        If tUser <= 0 Then
                            Call WriteConsoleMsg(UserIndex, "El usuario esta offline o no existe.", FontTypeNames.FONTTYPE_INFO)
                            Call LogGM(.Name, "Intento editar un usuario inexistente u offline.")
                        Else
                            Call ChangeUserChar(tUser, val(Arg1), UserList(tUser).Char.Head, UserList(tUser).Char.Heading, UserList(tUser).Char.WeaponAnim, UserList(tUser).Char.ShieldAnim, UserList(tUser).Char.CascoAnim, UserList(tUser).Char.AuraAnim, UserList(tUser).Char.AuraColor)

                        End If
                        
                        ' Log it
                        CommandString = CommandString & "BODY "
                    
                Case eEditOptions.eo_Head

                        If tUser <= 0 Then
                            Call WriteConsoleMsg(UserIndex, "El usuario esta offline o no existe.", FontTypeNames.FONTTYPE_INFO)
                            Call LogGM(.Name, "Intento editar un usuario inexistente u offline.")
                        Else
                            Call ChangeUserChar(tUser, UserList(tUser).Char.body, val(Arg1), UserList(tUser).Char.Heading, UserList(tUser).Char.WeaponAnim, UserList(tUser).Char.ShieldAnim, UserList(tUser).Char.CascoAnim, UserList(tUser).Char.AuraAnim, UserList(tUser).Char.AuraColor)

                        End If
                        
                        ' Log it
                        CommandString = CommandString & "HEAD "
                    
                Case eEditOptions.eo_CriminalsKilled
                        Var = IIf(val(Arg1) > MAXUSERMATADOS, MAXUSERMATADOS, val(Arg1))
                        
                        If tUser <= 0 Then ' Offline
                            Call WriteConsoleMsg(UserIndex, "El usuario esta offline o no existe.", FontTypeNames.FONTTYPE_INFO)
                            Call LogGM(.Name, "Intento editar un usuario inexistente u offline.")
                        Else ' Online
                            UserList(tUser).Faccion.CriminalesMatados = Var

                        End If
                        
                        ' Log it
                        CommandString = CommandString & "CRI "
                    
                Case eEditOptions.eo_CiticensKilled
                        Var = IIf(val(Arg1) > MAXUSERMATADOS, MAXUSERMATADOS, val(Arg1))
                        
                        If tUser <= 0 Then ' Offline
                            Call WriteConsoleMsg(UserIndex, "El usuario esta offline o no existe.", FontTypeNames.FONTTYPE_INFO)
                            Call LogGM(.Name, "Intento editar un usuario inexistente u offline.")
                        Else ' Online
                            UserList(tUser).Faccion.CiudadanosMatados = Var

                        End If
                        
                        ' Log it
                        CommandString = CommandString & "CIU "
                    
                Case eEditOptions.eo_Level

                        If val(Arg1) > STAT_MAXELV Then
                            Arg1 = CStr(STAT_MAXELV)
                            Call WriteConsoleMsg(UserIndex, "No puedes tener un nivel superior a " & STAT_MAXELV & ".", FONTTYPE_INFO)

                        End If
                        
                        ' Chequeamos si puede permanecer en el clan
                        If val(Arg1) >= 25 Then
                            
                            Dim GI As Integer

                            If tUser <= 0 Then
                                Call WriteConsoleMsg(UserIndex, "El usuario esta offline o no existe.", FontTypeNames.FONTTYPE_INFO)
                                Call LogGM(.Name, "Intento editar un usuario inexistente u offline.")
                            Else
                                GI = UserList(tUser).GuildIndex

                            End If
                            
                            If GI > 0 Then
                                If modGuilds.GuildAlignment(GI) = "Del Mal" Or modGuilds.GuildAlignment(GI) = "Real" Then
                                    'We get here, so guild has factionary alignment, we have to expulse the user
                                    Call modGuilds.m_EcharMiembroDeClan(-1, username)
                                    
                                    Call SendData(SendTarget.ToGuildMembers, GI, PrepareMessageConsoleMsg(username & " deja el clan.", FontTypeNames.FONTTYPE_GUILD))

                                    ' Si esta online le avisamos
                                    If tUser > 0 Then Call WriteConsoleMsg(tUser, "Ya tienes la madurez suficiente como para decidir bajo que estandarte pelearas! Por esta razon, hasta tanto no te enlistes en la faccion bajo la cual tu clan esta alineado, estaras excluido del mismo.", FontTypeNames.FONTTYPE_GUILD)

                                End If

                            End If

                        End If
                        
                        If tUser <= 0 Then ' Offline
                            Call WriteConsoleMsg(UserIndex, "El usuario esta offline o no existe.", FontTypeNames.FONTTYPE_INFO)
                            Call LogGM(.Name, "Intento editar un usuario inexistente u offline.")
                        Else ' Online
                            UserList(tUser).Stats.ELV = val(Arg1)
                            Call WriteUpdateUserStats(tUser)

                        End If
                    
                        ' Log it
                        CommandString = CommandString & "LEVEL "
                    
                Case eEditOptions.eo_Class

                        For LoopC = 1 To NUMCLASES

                            If UCase$(ListaClases(LoopC)) = UCase$(Arg1) Then Exit For
                        Next LoopC
                            
                        If LoopC > NUMCLASES Then
                            Call WriteConsoleMsg(UserIndex, "Clase desconocida. Intente nuevamente.", FontTypeNames.FONTTYPE_INFO)
                        Else

                            If tUser <= 0 Then ' Offline
                                Call WriteConsoleMsg(UserIndex, "El usuario esta offline o no existe.", FontTypeNames.FONTTYPE_INFO)
                                Call LogGM(.Name, "Intento editar un usuario inexistente u offline.")
                            Else ' Online
                                UserList(tUser).clase = LoopC

                            End If

                        End If
                    
                        ' Log it
                        CommandString = CommandString & "CLASE "
                        
                Case eEditOptions.eo_Skills

                        For LoopC = 1 To NUMSKILLS

                            If UCase$(Replace$(SkillsNames(LoopC), " ", "+")) = UCase$(Arg1) Then Exit For
                            
                        Next LoopC
                        
                        If LoopC > NUMSKILLS Then
                            Call WriteConsoleMsg(UserIndex, "Skill Inexistente!", FontTypeNames.FONTTYPE_INFO)
                        Else

                            If tUser <= 0 Then ' Offline
                                Call WriteConsoleMsg(UserIndex, "El usuario esta offline o no existe.", FontTypeNames.FONTTYPE_INFO)
                                Call LogGM(.Name, "Intento editar un usuario inexistente u offline.")
                            Else ' Online
                                UserList(tUser).Stats.UserSkills(LoopC) = val(Arg2)
                                Call CheckEluSkill(tUser, LoopC, True)

                            End If

                        End If
                        
                        ' Log it
                        CommandString = CommandString & "SKILLS "
                    
                Case eEditOptions.eo_Nobleza
                        Var = IIf(val(Arg1) > MAXREP, MAXREP, val(Arg1))
                        
                        If tUser <= 0 Then ' Offline
                            Call WriteConsoleMsg(UserIndex, "El usuario esta offline o no existe.", FontTypeNames.FONTTYPE_INFO)
                            Call LogGM(.Name, "Intento editar un usuario inexistente u offline.")
                        Else ' Online
                            UserList(tUser).Reputacion.NobleRep = Var

                        End If
                    
                        ' Log it
                        CommandString = CommandString & "NOB "
                        
                Case eEditOptions.eo_Asesino
                    Var = IIf(val(Arg1) > MAXREP, MAXREP, val(Arg1))
                        
                    If tUser <= 0 Then ' Offline
                        Call WriteConsoleMsg(UserIndex, "El usuario esta offline o no existe.", FontTypeNames.FONTTYPE_INFO)
                        Call LogGM(.Name, "Intento editar un usuario inexistente u offline.")
                    Else ' Online
                        UserList(tUser).Reputacion.AsesinoRep = Var

                    End If
                        
                    ' Log it
                    CommandString = CommandString & "ASE "
                    
                Case eEditOptions.eo_Sex

                    Dim Sex As Byte

                    Sex = IIf(UCase(Arg1) = "MUJER", eGenero.Mujer, 0) ' Mujer?
                    Sex = IIf(UCase(Arg1) = "HOMBRE", eGenero.Hombre, Sex) ' Hombre?
                        
                    If Sex <> 0 Then ' Es Hombre o mujer?
                        If tUser <= 0 Then ' OffLine
                            Call WriteConsoleMsg(UserIndex, "El usuario esta offline o no existe.", FontTypeNames.FONTTYPE_INFO)
                            Call LogGM(.Name, "Intento editar un usuario inexistente u offline.")
                        Else ' Online
                            UserList(tUser).Genero = Sex

                        End If

                    Else
                        Call WriteConsoleMsg(UserIndex, "Genero desconocido. Intente nuevamente.", FontTypeNames.FONTTYPE_INFO)

                    End If
                        
                    ' Log it
                    CommandString = CommandString & "SEX "
                    
                Case eEditOptions.eo_Raza

                    Dim Raza As Byte
                        
                    Arg1 = UCase$(Arg1)

                    Select Case Arg1

                        Case "HUMANO"
                            Raza = eRaza.Humano

                        Case "ELFO"
                            Raza = eRaza.Elfo

                        Case "DROW"
                            Raza = eRaza.Drow

                        Case "ENANO"
                            Raza = eRaza.Enano

                        Case "GNOMO"
                            Raza = eRaza.Gnomo
                                
                        Case "ORCO"
                            Raza = eRaza.Orco

                        Case "VAMPIRO"
                            Raza = eRaza.Vampiro
                            
                        Case Else
                            Raza = 0

                        End Select
                            
                    If Raza = 0 Then
                        Call WriteConsoleMsg(UserIndex, "Raza desconocida. Intente nuevamente.", FontTypeNames.FONTTYPE_INFO)
                    Else

                        If tUser <= 0 Then
                            Call WriteConsoleMsg(UserIndex, "El usuario esta offline o no existe.", FontTypeNames.FONTTYPE_INFO)
                            Call LogGM(.Name, "Intento editar un usuario inexistente u offline.")
                        Else
                            UserList(tUser).Raza = Raza

                        End If

                    End If
                            
                    ' Log it
                    CommandString = CommandString & "RAZA "
                        
                Case eEditOptions.eo_addGold
                    
                    Dim bankGold As Long
                        
                    If Abs(Arg1) > MAX_ORO_EDIT Then
                        Call WriteConsoleMsg(UserIndex, "No esta permitido utilizar valores mayores a " & MAX_ORO_EDIT & ".", FontTypeNames.FONTTYPE_INFO)
                    Else

                        If tUser <= 0 Then
                            Call WriteConsoleMsg(UserIndex, "El usuario esta offline o no existe.", FontTypeNames.FONTTYPE_INFO)
                            Call LogGM(.Name, "Intento editar un usuario inexistente u offline.")

                        End If

                    End If
                        
                    ' Log it
                    CommandString = CommandString & "AGREGAR "
                    
                Case eEditOptions.eo_Vida
                    
                    If val(Arg1) > MAX_VIDA_EDIT Then
                        Arg1 = CStr(MAX_VIDA_EDIT)
                        Call WriteConsoleMsg(UserIndex, "No puedes tener vida superior a " & MAX_VIDA_EDIT & ".", FONTTYPE_INFO)

                    End If
                        
                    ' No valido si esta offline, porque solo se puede editar a si mismo
                    UserList(tUser).Stats.MaxHp = val(Arg1)
                    UserList(tUser).Stats.MinHp = val(Arg1)
                        
                    Call WriteUpdateUserStats(tUser)
                        
                    ' Log it
                    CommandString = CommandString & "VIDA "
                        
                Case eEditOptions.eo_Poss
                    
                    Dim Map As Integer

                    Dim X   As Integer

                    Dim Y   As Integer
                        
                    Map = val(ReadField(1, Arg1, 45))
                    X = val(ReadField(2, Arg1, 45))
                    Y = val(ReadField(3, Arg1, 45))
                        
                    If InMapBounds(Map, X, Y) Then
                            
                        If tUser <= 0 Then
                            Call WriteConsoleMsg(UserIndex, "El usuario esta offline o no existe.", FontTypeNames.FONTTYPE_INFO)
                            Call LogGM(.Name, "Intento editar un usuario inexistente u offline.")
                        Else
                            Call WarpUserChar(tUser, Map, X, Y, True, True)
                            Call WriteConsoleMsg(UserIndex, "Usuario teletransportado: " & username, FontTypeNames.FONTTYPE_INFO)

                        End If

                    Else
                        Call WriteConsoleMsg(UserIndex, "Posicion invalida", FONTTYPE_INFO)

                    End If
                        
                    ' Log it
                    CommandString = CommandString & "POSS "
                    
                    Case eEditOptions.eo_Speed
                        
                        Dim Speed As Double
                        
                        If val(Arg1) > 50 Then _
                            Arg1 = 50
                            
                        Speed = val(Arg1)
                        
                        UserList(tUser).flags.Velocidad = Speed
                        Call WriteSetSpeed(tUser)
                        Call SendData(SendTarget.ToPCArea, tUser, PrepareMessageSpeeding(UserList(tUser).Char.CharIndex, UserList(tUser).flags.Velocidad))
                        
                Case Else
                    Call WriteConsoleMsg(UserIndex, "Comando no permitido.", FontTypeNames.FONTTYPE_INFO)
                    CommandString = CommandString & "UNKOWN "
                        
            End Select
                
            CommandString = CommandString & Arg1 & " " & Arg2
                
            If UserIndex <> tUser Then
                Call LogGM(.Name, CommandString & " " & username)
            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
        
    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "RequestCharInfo" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestCharInfo(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Fredy Horacio Treboux (liquid)
    'Last Modification: 01/08/07
    'Last Modification by: (liquid).. alto bug zapallo..
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
                
        Dim TargetName  As String

        Dim targetIndex As Integer
        
        TargetName = Replace$(Buffer.ReadASCIIString(), "+", " ")
        targetIndex = NameIndex(TargetName)
        
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios) Then

            'is the player offline?
            If targetIndex <= 0 Then

                'don't allow to retrieve administrator's info
                If Not (EsDios(TargetName) Or EsAdmin(TargetName)) Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline, buscando...", FontTypeNames.FONTTYPE_INFO)

                    Call SendUserStatsTxtDatabase(UserIndex, TargetName)

                End If

            Else

                'don't allow to retrieve administrator's info
                If UserList(targetIndex).flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then
                    Call SendUserStatsTxt(UserIndex, targetIndex)

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "RequestCharStats" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestCharStats(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 07/06/2010
    '07/06/2010: ZaMa - Ahora no se puede usar para saber si hay dioses/admins online.
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim username         As String

        Dim tUser            As Integer
        
        Dim UserIsAdmin      As Boolean

        Dim OtherUserIsAdmin As Boolean
        
        username = Buffer.ReadASCIIString()
         
        UserIsAdmin = (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And ((.flags.Privilegios And PlayerType.SemiDios) <> 0 Or UserIsAdmin) Then
            Call LogGM(.Name, "/STAT " & username)
            
            tUser = NameIndex(username)
            
            OtherUserIsAdmin = EsDios(username) Or EsAdmin(username)
            
            If tUser <= 0 Then
                If UserIsAdmin Or Not OtherUserIsAdmin Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline. Buscando... ", FontTypeNames.FONTTYPE_INFO)

                    Call SendUserMiniStatsTxtFromDatabase(UserIndex, username)

                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes ver los stats de un dios o admin.", FontTypeNames.FONTTYPE_INFO)

                End If

            Else

                If UserIsAdmin Or Not OtherUserIsAdmin Then
                    Call SendUserMiniStatsTxt(UserIndex, tUser)
                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes ver los stats de un dios o admin.", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "RequestCharGold" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestCharGold(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 07/06/2010
    '07/06/2010: ZaMa - Ahora no se puede usar para saber si hay dioses/admins online.
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim username         As String

        Dim tUser            As Integer
        
        Dim UserIsAdmin      As Boolean

        Dim OtherUserIsAdmin As Boolean
        
        username = Buffer.ReadASCIIString()
        
        UserIsAdmin = (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0
        
        If (.flags.Privilegios And PlayerType.SemiDios) Or UserIsAdmin Then
            
            Call LogGM(.Name, "/BAL " & username)
            
            tUser = NameIndex(username)
            OtherUserIsAdmin = EsDios(username) Or EsAdmin(username)
            
            tUser = NameIndex(username)
            OtherUserIsAdmin = EsDios(username) Or EsAdmin(username)
            
            If tUser <= 0 Then
                If UserIsAdmin Or Not OtherUserIsAdmin Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline. Buscando... ", FontTypeNames.FONTTYPE_TALK)

                    Call SendUserOROTxtFromDatabase(UserIndex, username)

                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes ver el oro de un dios o admin.", FontTypeNames.FONTTYPE_INFO)

                End If

            Else

                If UserIsAdmin Or Not OtherUserIsAdmin Then
                    Call WriteConsoleMsg(UserIndex, "El usuario " & username & " tiene " & UserList(tUser).Stats.Banco & " en el banco.", FontTypeNames.FONTTYPE_TALK)
                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes ver el oro de un dios o admin.", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "RequestCharInventory" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestCharInventory(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 07/06/2010
    '07/06/2010: ZaMa - Ahora no se puede usar para saber si hay dioses/admins online.
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim username         As String

        Dim tUser            As Integer
        
        Dim UserIsAdmin      As Boolean

        Dim OtherUserIsAdmin As Boolean
        
        username = Buffer.ReadASCIIString()
        
        UserIsAdmin = (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            Call LogGM(.Name, "/INV " & username)
            
            tUser = NameIndex(username)
            OtherUserIsAdmin = EsDios(username) Or EsAdmin(username)
            
            tUser = NameIndex(username)
            OtherUserIsAdmin = EsDios(username) Or EsAdmin(username)
            
            If tUser <= 0 Then
                If UserIsAdmin Or Not OtherUserIsAdmin Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline. Buscando...", FontTypeNames.FONTTYPE_TALK)

                    Call SendUserInvTxtFromDatabase(UserIndex, username)

                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes ver el inventario de un dios o admin.", FontTypeNames.FONTTYPE_INFO)

                End If

            Else

                If UserIsAdmin Or Not OtherUserIsAdmin Then
                    Call SendUserInvTxt(UserIndex, tUser)
                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes ver el inventario de un dios o admin.", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "RequestCharBank" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestCharBank(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 07/06/2010
    '07/06/2010: ZaMa - Ahora no se puede usar para saber si hay dioses/admins online.
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim username         As String

        Dim tUser            As Integer
        
        Dim UserIsAdmin      As Boolean

        Dim OtherUserIsAdmin As Boolean

        username = Buffer.ReadASCIIString()
        
        UserIsAdmin = (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0
        
        If (.flags.Privilegios And PlayerType.SemiDios) <> 0 Or UserIsAdmin Then
            Call LogGM(.Name, "/BOV " & username)
            
            tUser = NameIndex(username)
            OtherUserIsAdmin = EsDios(username) Or EsAdmin(username)
            
            tUser = NameIndex(username)
            OtherUserIsAdmin = EsDios(username) Or EsAdmin(username)
            
            If tUser <= 0 Then
                If UserIsAdmin Or Not OtherUserIsAdmin Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline. Buscando... ", FontTypeNames.FONTTYPE_TALK)

                    Call SendUserBovedaTxtFromDatabase(UserIndex, username)

                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes ver la boveda de un dios o admin.", FontTypeNames.FONTTYPE_INFO)

                End If

            Else

                If UserIsAdmin Or Not OtherUserIsAdmin Then
                    Call SendUserBovedaTxt(UserIndex, tUser)
                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes ver la boveda de un dios o admin.", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "RequestCharSkills" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRequestCharSkills(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim username As String

        Dim tUser    As Integer

        Dim LoopC    As Long

        Dim Message  As String
        
        username = Buffer.ReadASCIIString()
        tUser = NameIndex(username)
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            Call LogGM(.Name, "/STATS " & username)
            
            If tUser <= 0 Then
                If (InStrB(username, "\") <> 0) Then
                    username = Replace(username, "\", "")

                End If

                If (InStrB(username, "/") <> 0) Then
                    username = Replace(username, "/", "")

                End If
                
                For LoopC = 1 To NUMSKILLS
                    Message = Message & GetUserSkills(username)
                Next LoopC
                
                Call WriteConsoleMsg(UserIndex, Message & "CHAR> Libres: " & GetUserFreeSkills(username), FontTypeNames.FONTTYPE_INFO)

            Else
                Call SendUserSkillsTxt(UserIndex, tUser)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "ReviveChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleReviveChar(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 11/03/2010
    '11/03/2010: ZaMa - Al revivir con el comando, si esta navegando le da cuerpo e barca.
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim username As String

        Dim tUser    As Integer

        Dim LoopC    As Byte
        
        username = Buffer.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            If UCase$(username) <> "YO" Then
                tUser = NameIndex(username)
            Else
                tUser = UserIndex

            End If
            
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
            Else

                With UserList(tUser)

                    'If dead, show him alive (naked).
                    If .flags.Muerto = 1 Then
                        .flags.Muerto = 0
                        
                        If .flags.Navegando = 1 Then
                            Call ToggleBoatBody(tUser)
                        Else
                            Call DarCuerpoDesnudo(tUser)

                        End If
                        
                        Call ChangeUserChar(tUser, .Char.body, .OrigChar.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraAnim, .Char.AuraColor)
                        Call UpdateUserSpeed(tUser)
                        
                        Call WriteConsoleMsg(tUser, UserList(UserIndex).Name & " te ha resucitado.", FontTypeNames.FONTTYPE_INFO)
                    Else
                        Call WriteConsoleMsg(tUser, UserList(UserIndex).Name & " te ha curado.", FontTypeNames.FONTTYPE_INFO)

                    End If
                    
                    .Stats.MinHp = .Stats.MaxHp
                    
                End With
                
                Call WriteUpdateHP(tUser)
                
                Call LogGM(.Name, "Resucito a " & username)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "OnlineGM" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleOnlineGM(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Fredy Horacio Treboux (liquid)
    'Last Modification: 12/28/06
    '
    '***************************************************
    Dim i    As Long

    Dim list As String

    Dim priv As PlayerType
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub

        priv = PlayerType.Consejero Or PlayerType.SemiDios

        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then priv = priv Or PlayerType.Dios Or PlayerType.Admin
        
        For i = 1 To LastUser

            If UserList(i).flags.UserLogged Then
                If UserList(i).flags.Privilegios And priv Then list = list & UserList(i).Name & ", "

            End If

        Next i
        
        If LenB(list) <> 0 Then
            list = Left$(list, Len(list) - 2)
            Call WriteConsoleMsg(UserIndex, list & ".", FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(UserIndex, "No hay GMs Online.", FontTypeNames.FONTTYPE_INFO)

        End If

    End With

End Sub

''
' Handles the "OnlineMap" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleOnlineMap(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 23/03/2009
    '23/03/2009: ZaMa - Ahora no requiere estar en el mapa, sino que por defecto se toma en el que esta, pero se puede especificar otro
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Map As Integer

        Map = .incomingData.ReadInteger
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub
        
        Dim LoopC As Long

        Dim list  As String

        Dim priv  As PlayerType
        
        priv = PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios

        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then priv = priv + (PlayerType.Dios Or PlayerType.Admin)
        
        For LoopC = 1 To LastUser

            If LenB(UserList(LoopC).Name) <> 0 And UserList(LoopC).Pos.Map = Map Then
                If UserList(LoopC).flags.Privilegios And priv Then list = list & UserList(LoopC).Name & ", "

            End If

        Next LoopC
        
        If Len(list) > 2 Then list = Left$(list, Len(list) - 2)
        
        Call WriteConsoleMsg(UserIndex, "Usuarios en el mapa: " & list, FontTypeNames.FONTTYPE_INFO)
        Call LogGM(.Name, "/ONLINEMAP " & Map)

    End With

End Sub

''
' Handles the "Forgive" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleForgive(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 07/06/2010
    '07/06/2010: ZaMa - Ahora no se puede usar para saber si hay dioses/admins online.
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim username As String

        Dim tUser    As Integer
        
        username = Buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            tUser = NameIndex(username)
            
            If tUser > 0 Then
                If EsNewbie(tUser) Then
                    Call VolverCiudadano(tUser)
                Else
                    Call LogGM(.Name, "Intento perdonar un personaje de nivel avanzado.")
                    
                    If Not (EsDios(username) Or EsAdmin(username)) Then
                        Call WriteConsoleMsg(UserIndex, "Solo se permite perdonar newbies.", FontTypeNames.FONTTYPE_INFO)

                    End If

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "Kick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleKick(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 07/06/2010
    '07/06/2010: ZaMa - Ahora no se puede usar para saber si hay dioses/admins online.
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim username As String

        Dim tUser    As Integer

        Dim Rank     As Integer

        Dim IsAdmin  As Boolean
        
        Rank = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero
        
        username = Buffer.ReadASCIIString()
        IsAdmin = (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0
        
        If (.flags.Privilegios And PlayerType.SemiDios) Or IsAdmin Then
            tUser = NameIndex(username)
            
            If tUser <= 0 Then
                If Not (EsDios(username) Or EsAdmin(username)) Or IsAdmin Then
                    Call WriteConsoleMsg(UserIndex, "El usuario no esta online.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes echar a alguien con jerarquia mayor a la tuya.", FontTypeNames.FONTTYPE_INFO)

                End If

            Else

                If (UserList(tUser).flags.Privilegios And Rank) > (.flags.Privilegios And Rank) Then
                    Call WriteConsoleMsg(UserIndex, "No puedes echar a alguien con jerarquia mayor a la tuya.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call SendData(SendTarget.Toall, 0, PrepareMessageConsoleMsg(.Name & " echo a " & username & ".", FontTypeNames.FONTTYPE_INFO))
                    Call CloseUser(tUser)
                    Call LogGM(.Name, "Echo a " & username)

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "Execute" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleExecute(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 07/06/2010
    '07/06/2010: ZaMa - Ahora no se puede usar para saber si hay dioses/admins online.
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim username As String

        Dim tUser    As Integer
        
        username = Buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            tUser = NameIndex(username)
            
            If tUser > 0 Then
                If Not UserList(tUser).flags.Privilegios And PlayerType.User Then
                    Call WriteConsoleMsg(UserIndex, "Estas loco?? Como vas a pinatear un gm?? :@", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call UserDie(tUser)
                    Call SendData(SendTarget.Toall, 0, PrepareMessageConsoleMsg(.Name & " ha ejecutado a " & username & ".", FontTypeNames.FONTTYPE_EJECUCION))
                    Call LogGM(.Name, " ejecuto a " & username)

                End If

            Else

                If Not (EsDios(username) Or EsAdmin(username)) Then
                    Call WriteConsoleMsg(UserIndex, "No esta online.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, "Estas loco?? Como vas a pinatear un gm?? :@", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "BanChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBanChar(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim username As String

        Dim Reason   As String
        
        username = Buffer.ReadASCIIString()
        Reason = Buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            Call BanCharacter(UserIndex, username, Reason)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "UnbanChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUnbanChar(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim username  As String

        Dim cantPenas As Byte
        
        username = Buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            If (InStrB(username, "\") <> 0) Then
                username = Replace(username, "\", "")

            End If

            If (InStrB(username, "/") <> 0) Then
                username = Replace(username, "/", "")

            End If
            
            If Not PersonajeExiste(username) Then
                Call WriteConsoleMsg(UserIndex, "Charfile inexistente (no use +).", FontTypeNames.FONTTYPE_INFO)
            Else

                If BANCheck(username) Then
                    Call UnBan(username)
                
                    'penas
                    cantPenas = GetUserAmountOfPunishments(username)
                    Call SaveUserPunishment(username, cantPenas + 1, LCase$(.Name) & ": UNBAN. " & Date & " " & time)
                
                    Call LogGM(.Name, "/UNBAN a " & username)
                    Call WriteConsoleMsg(UserIndex, username & " unbanned.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, username & " no esta baneado. Imposible unbanear.", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "NPCFollow" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleNPCFollow(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub
        
        If .flags.TargetNPC > 0 Then
            Call DoFollow(.flags.TargetNPC, .Name)
            Npclist(.flags.TargetNPC).flags.Inmovilizado = 0
            Npclist(.flags.TargetNPC).flags.Paralizado = 0
            Npclist(.flags.TargetNPC).Contadores.Paralisis = 0

        End If

    End With

End Sub

''
' Handles the "SummonChar" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSummonChar(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 26/03/2009
    '26/03/2009: ZaMa - Chequeo que no se teletransporte donde haya un char o npc
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim username As String

        Dim tUser    As Integer

        Dim X        As Integer

        Dim Y        As Integer
        
        username = Buffer.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            tUser = NameIndex(username)
            
            If tUser <= 0 Then
                If EsDios(username) Or EsAdmin(username) Then
                    Call WriteConsoleMsg(UserIndex, "No puedes invocar a dioses y admins.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, "El jugador no esta online.", FontTypeNames.FONTTYPE_INFO)

                End If
                
            Else

                If (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0 Or (UserList(tUser).flags.Privilegios And (PlayerType.Consejero Or PlayerType.User)) <> 0 Then
                    Call WriteConsoleMsg(tUser, .Name & " te ha trasportado.", FontTypeNames.FONTTYPE_INFO)
                    X = .Pos.X
                    Y = .Pos.Y + 1
                    Call FindLegalPos(tUser, .Pos.Map, X, Y)
                    Call WarpUserChar(tUser, .Pos.Map, X, Y, True, True)
                    Call LogGM(.Name, "/SUM " & username & " Map:" & .Pos.Map & " X:" & .Pos.X & " Y:" & .Pos.Y)
                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes invocar a dioses y admins.", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "SpawnListRequest" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSpawnListRequest(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub
        
        Call EnviarSpawnList(UserIndex)

    End With

End Sub

''
' Handles the "SpawnCreature" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSpawnCreature(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim NPC As Integer

        NPC = .incomingData.ReadInteger()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            If NPC > 0 And NPC <= UBound(Declaraciones.SpawnList()) Then Call SpawnNpc(Declaraciones.SpawnList(NPC).NPCIndex, .Pos, True, False)
            
            Call LogGM(.Name, "Sumoneo " & Declaraciones.SpawnList(NPC).NpcName)

        End If

    End With

End Sub

''
' Handles the "ResetNPCInventory" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleResetNPCInventory(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster) Then Exit Sub
        If .flags.TargetNPC = 0 Then Exit Sub
        
        Call ResetNpcInv(.flags.TargetNPC)
        Call LogGM(.Name, "/RESETINV " & Npclist(.flags.TargetNPC).Name)

    End With

End Sub

''
' Handles the "ServerMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleServerMessage(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 28/05/2010
    '28/05/2010: ZaMa - Ahora no dice el nombre del gm que lo dice.
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim Message As String

        Message = Buffer.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            If LenB(Message) <> 0 Then
                Call LogGM(.Name, "Mensaje Broadcast:" & Message)
                Call SendData(SendTarget.Toall, 0, PrepareMessageConsoleMsg(Message, FontTypeNames.FONTTYPE_TALK))

                ''''''''''''''''SOLO PARA EL TESTEO'''''''
                ''''''''''SE USA PARA COMUNICARSE CON EL SERVER'''''''''''
                'frmMain.txtChat.Text = frmMain.txtChat.Text & vbNewLine & UserList(UserIndex).name & " > " & message
            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "MapMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleMapMessage(ByVal UserIndex As Integer)

    '***************************************************
    'Author: ZaMa
    'Last Modification: 14/11/2010
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim Message As String

        Message = Buffer.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) Then
            If LenB(Message) <> 0 Then
                
                Dim Mapa As Integer
                                        Mapa = .Pos.Map

                Call LogGM(.Name, "Mensaje a mapa " & Mapa & ":" & Message)
                Call SendData(SendTarget.toMap, Mapa, PrepareMessageConsoleMsg(Message, FontTypeNames.FONTTYPE_TALK))

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "NickToIP" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleNickToIP(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 07/06/2010
    'Pablo (ToxicWaste): Agrego para que el /nick2ip tambien diga los nicks en esa ip por pedido de la DGM.
    '07/06/2010: ZaMa - Ahora no se puede usar para saber si hay dioses/admins online.
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim username As String

        Dim tUser    As Integer

        Dim priv     As PlayerType

        Dim IsAdmin  As Boolean
        
        username = Buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            tUser = NameIndex(username)
            Call LogGM(.Name, "NICK2IP Solicito la IP de " & username)
            
            IsAdmin = (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0

            If IsAdmin Then
                priv = PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.Dios Or PlayerType.Admin
            Else
                priv = PlayerType.User

            End If
            
            If tUser > 0 Then
                If UserList(tUser).flags.Privilegios And priv Then
                    Call WriteConsoleMsg(UserIndex, "El ip de " & username & " es " & UserList(tUser).IP, FontTypeNames.FONTTYPE_INFO)

                    Dim IP    As String

                    Dim lista As String

                    Dim LoopC As Long

                    IP = UserList(tUser).IP

                    For LoopC = 1 To LastUser

                        If UserList(LoopC).IP = IP Then
                            If LenB(UserList(LoopC).Name) <> 0 And UserList(LoopC).flags.UserLogged Then
                                If UserList(LoopC).flags.Privilegios And priv Then
                                    lista = lista & UserList(LoopC).Name & ", "

                                End If

                            End If

                        End If

                    Next LoopC

                    If LenB(lista) <> 0 Then lista = Left$(lista, Len(lista) - 2)
                    Call WriteConsoleMsg(UserIndex, "Los personajes con ip " & IP & " son: " & lista, FontTypeNames.FONTTYPE_INFO)

                End If

            Else

                If Not (EsDios(username) Or EsAdmin(username)) Or IsAdmin Then
                    Call WriteConsoleMsg(UserIndex, "No hay ningUn personaje con ese nick.", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "IPToNick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleIPToNick(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim IP    As String

        Dim LoopC As Long

        Dim lista As String

        Dim priv  As PlayerType
        
        IP = .incomingData.ReadByte() & "."
        IP = IP & .incomingData.ReadByte() & "."
        IP = IP & .incomingData.ReadByte() & "."
        IP = IP & .incomingData.ReadByte()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, "IP2NICK Solicito los Nicks de IP " & IP)
        
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) Then
            priv = PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.Dios Or PlayerType.Admin
        Else
            priv = PlayerType.User

        End If

        For LoopC = 1 To LastUser

            If UserList(LoopC).IP = IP Then
                If LenB(UserList(LoopC).Name) <> 0 And UserList(LoopC).flags.UserLogged Then
                    If UserList(LoopC).flags.Privilegios And priv Then
                        lista = lista & UserList(LoopC).Name & ", "

                    End If

                End If

            End If

        Next LoopC
        
        If LenB(lista) <> 0 Then lista = Left$(lista, Len(lista) - 2)
        Call WriteConsoleMsg(UserIndex, "Los personajes con ip " & IP & " son: " & lista, FontTypeNames.FONTTYPE_INFO)

    End With

End Sub

''
' Handles the "GuildOnlineMembers" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildOnlineMembers(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim GuildName As String

        Dim tGuild    As Integer
        
        GuildName = Buffer.ReadASCIIString()
        
        If (InStrB(GuildName, "+") <> 0) Then
            GuildName = Replace(GuildName, "+", " ")

        End If
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            tGuild = GuildIndex(GuildName)
            
            If tGuild > 0 Then
                Call WriteConsoleMsg(UserIndex, "Clan " & UCase(GuildName) & ": " & modGuilds.m_ListaDeMiembrosOnline(UserIndex, tGuild), FontTypeNames.FONTTYPE_GUILDMSG)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "TeleportCreate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTeleportCreate(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 22/03/2010
    '15/11/2009: ZaMa - Ahora se crea un teleport con un radio especificado.
    '22/03/2010: ZaMa - Harcodeo los teleps y radios en el dat, para evitar mapas bugueados.
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 6 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim Mapa  As Integer

        Dim X     As Integer

        Dim Y     As Integer

        Dim Radio As Byte
        
        Mapa = .incomingData.ReadInteger()
        X = .incomingData.ReadInteger()
        Y = .incomingData.ReadInteger()
        Radio = .incomingData.ReadByte()
        
        Radio = MinimoInt(Radio, 6)
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Call LogGM(.Name, "/CT " & Mapa & "," & X & "," & Y & "," & Radio)
        
        If Not MapaValido(Mapa) Or Not InMapBounds(Mapa, X, Y) Then Exit Sub
        
        If MapData(.Pos.Map, .Pos.X, .Pos.Y - 1).ObjInfo.ObjIndex > 0 Then Exit Sub
        
        If MapData(.Pos.Map, .Pos.X, .Pos.Y - 1).TileExit.Map > 0 Then Exit Sub
        
        If MapData(Mapa, X, Y).ObjInfo.ObjIndex > 0 Then
            Call WriteConsoleMsg(UserIndex, "Hay un objeto en el piso en ese lugar.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        If MapData(Mapa, X, Y).TileExit.Map > 0 Then
            Call WriteConsoleMsg(UserIndex, "No puedes crear un teleport que apunte a la entrada de otro.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        Dim ET As obj

        ET.Amount = 1
        ' Es el numero en el dat. El indice es el comienzo + el radio, todo harcodeado :(.
        ET.ObjIndex = TELEP_OBJ_INDEX + Radio
        
        With MapData(.Pos.Map, .Pos.X, .Pos.Y - 1)
            .TileExit.Map = Mapa
            .TileExit.X = X
            .TileExit.Y = Y

        End With
        
        Call MakeObj(ET, .Pos.Map, .Pos.X, .Pos.Y - 1)

    End With

End Sub

''
' Handles the "TeleportDestroy" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTeleportDestroy(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    With UserList(UserIndex)

        Dim Mapa As Integer

        Dim X    As Byte

        Dim Y    As Byte
        
        'Remove packet ID
        Call .incomingData.ReadByte
        
        '/dt
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Mapa = .flags.TargetMap
        X = .flags.TargetX
        Y = .flags.TargetY
        
        If Not InMapBounds(Mapa, X, Y) Then Exit Sub
        
        With MapData(Mapa, X, Y)

            If .ObjInfo.ObjIndex = 0 Then Exit Sub
            
            If ObjData(.ObjInfo.ObjIndex).OBJType = eOBJType.otTeleport And .TileExit.Map > 0 Then
                
                                Call LogGM(UserList(UserIndex).Name, "/DT: " & Mapa & "," & X & "," & Y)
                
                Call EraseObj(.ObjInfo.Amount, Mapa, X, Y)
                
                If MapData(.TileExit.Map, .TileExit.X, .TileExit.Y).ObjInfo.ObjIndex = 651 Then
                    Call EraseObj(1, .TileExit.Map, .TileExit.X, .TileExit.Y)

                End If
                
                .TileExit.Map = 0
                .TileExit.X = 0
                .TileExit.Y = 0

            End If

        End With

    End With

End Sub

''
' Handles the "ExitDestroy" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleExitDestroy(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Cucsifae
    'Last Modification: 30/9/18
    '
    '***************************************************
    With UserList(UserIndex)

        Dim Mapa As Integer

        Dim X    As Byte

        Dim Y    As Byte
        
        'Remove packet ID
        Call .incomingData.ReadByte
        
        '/de
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Mapa = .flags.TargetMap
        X = .flags.TargetX
        Y = .flags.TargetY
        
        If Not InMapBounds(Mapa, X, Y) Then Exit Sub
        
        With MapData(Mapa, X, Y)

            If .TileExit.Map = 0 Then Exit Sub

            'Si hay un Teleport hay que usar /DT
            If .ObjInfo.ObjIndex > 0 Then
                If ObjData(.ObjInfo.ObjIndex).OBJType = eOBJType.otTeleport Then Exit Sub

            End If

            Call LogGM(UserList(UserIndex).Name, "/DE: " & Mapa & "," & X & "," & Y)
                
            .TileExit.Map = 0
            .TileExit.X = 0
            .TileExit.Y = 0

        End With

    End With

End Sub

''
' Handles the "MeteoToggle" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleMeteoToggle(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    Dim Forzar As Byte
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Forzar = .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub
        
        Call LogGM(.Name, "/METEO " & Forzar)
        
        Lloviendo = Not Lloviendo
        
        Call SortearClima(Forzar)

    End With

End Sub

''
' Handles the "EnableDenounces" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleEnableDenounces(ByVal UserIndex As Integer)
    '***************************************************
    'Author: ZaMa
    'Last Modification: 14/11/2010
    'Enables/Disables
    '***************************************************

    With UserList(UserIndex)
    
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If Not EsGm(UserIndex) Then Exit Sub
        
        Dim Activado As Boolean

        Dim msg      As String
        
        Activado = Not .flags.SendDenounces
        .flags.SendDenounces = Activado
        
        msg = "Denuncias por consola " & IIf(Activado, "ativadas", "desactivadas") & "."
        
        Call LogGM(.Name, msg)
        
        Call WriteConsoleMsg(UserIndex, msg, FontTypeNames.FONTTYPE_INFO)

    End With

End Sub

''
' Handles the "ShowDenouncesList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleShowDenouncesList(ByVal UserIndex As Integer)

    '***************************************************
    'Author: ZaMa
    'Last Modification: 14/11/2010
    '
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And PlayerType.User Then Exit Sub
        Call WriteShowDenounces(UserIndex)

    End With

End Sub

''
' Handles the "SetDialog" message.
'
' @param UserIndex The index of the user sending the message

Public Sub HandleSetDialog(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Amraphen
    'Last Modification: 18/11/2010
    '20/11/2010: ZaMa - Arreglo privilegios.
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet id
        Call Buffer.ReadByte
        
        Dim NewDialog As String

        NewDialog = Buffer.ReadASCIIString
        
        Call .incomingData.CopyBuffer(Buffer)
        
        If .flags.TargetNPC > 0 Then

            ' Dsgm/Dsrm/Rm
            If Not ((.flags.Privilegios And PlayerType.Dios) = 0 And (.flags.Privilegios And (PlayerType.SemiDios Or PlayerType.RoleMaster)) <> (PlayerType.SemiDios Or PlayerType.RoleMaster)) Then
                'Replace the NPC's dialog.
                Npclist(.flags.TargetNPC).Desc = NewDialog

            End If

        End If

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "Impersonate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleImpersonate(ByVal UserIndex As Integer)

    '***************************************************
    'Author: ZaMa
    'Last Modification: 20/11/2010
    '
    '***************************************************
    With UserList(UserIndex)
    
        'Remove packet ID
        Call .incomingData.ReadByte
        
        ' Dsgm/Dsrm/Rm
        If (.flags.Privilegios And PlayerType.Dios) = 0 And (.flags.Privilegios And (PlayerType.SemiDios Or PlayerType.RoleMaster)) <> (PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Dim NPCIndex As Integer

        NPCIndex = .flags.TargetNPC
        
        If NPCIndex = 0 Then Exit Sub
        
        ' Copy head, body and desc
        Call ImitateNpc(UserIndex, NPCIndex)
        
        ' Teleports user to npc's coords
        Call WarpUserChar(UserIndex, Npclist(NPCIndex).Pos.Map, Npclist(NPCIndex).Pos.X, Npclist(NPCIndex).Pos.Y, False, True)
        
        ' Log gm
        Call LogGM(.Name, "/IMPERSONAR con " & Npclist(NPCIndex).Name & " en mapa " & .Pos.Map)
        
        ' Remove npc
        Call QuitarNPC(NPCIndex)
        
    End With
    
End Sub

''
' Handles the "Imitate" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleImitate(ByVal UserIndex As Integer)

    '***************************************************
    'Author: ZaMa
    'Last Modification: 20/11/2010
    '
    '***************************************************
    With UserList(UserIndex)
    
        'Remove packet ID
        Call .incomingData.ReadByte
        
        ' Dsgm/Dsrm/Rm/ConseRm
        If (.flags.Privilegios And PlayerType.Dios) = 0 And (.flags.Privilegios And (PlayerType.SemiDios Or PlayerType.RoleMaster)) <> (PlayerType.SemiDios Or PlayerType.RoleMaster) And (.flags.Privilegios And (PlayerType.Consejero Or PlayerType.RoleMaster)) <> (PlayerType.Consejero Or PlayerType.RoleMaster) Then Exit Sub
        
        Dim NPCIndex As Integer

        NPCIndex = .flags.TargetNPC
        
        If NPCIndex = 0 Then Exit Sub
        
        ' Copy head, body and desc
        Call ImitateNpc(UserIndex, NPCIndex)
        Call LogGM(.Name, "/MIMETIZAR con " & Npclist(NPCIndex).Name & " en mapa " & .Pos.Map)
        
    End With
    
End Sub

''
' Handles the "RecordAdd" message.
'
' @param UserIndex The index of the user sending the message
           
Public Sub HandleRecordAdd(ByVal UserIndex As Integer)

    '**************************************************************
    'Author: Amraphen
    'Last Modify Date: 29/11/2010
    '
    '**************************************************************
    If UserList(UserIndex).incomingData.Length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet id
        Call Buffer.ReadByte
        
        Dim username As String

        Dim Reason   As String
        
        username = Buffer.ReadASCIIString
        Reason = Buffer.ReadASCIIString
    
        If Not (.flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster)) Then

            'Verificamos que exista el personaje
            If Not PersonajeExiste(username) Then
                Call WriteShowMessageBox(UserIndex, "El personaje no existe")
            Else
                'Agregamos el seguimiento
                Call AddRecord(UserIndex, username, Reason)
                
                'Enviamos la nueva lista de personajes
                Call WriteRecordList(UserIndex)

            End If

        End If

        Call .incomingData.CopyBuffer(Buffer)

    End With
        
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "RecordAddObs" message.
'
' @param UserIndex The index of the user sending the message.

Public Sub HandleRecordAddObs(ByVal UserIndex As Integer)

    '**************************************************************
    'Author: Amraphen
    'Last Modify Date: 29/11/2010
    '
    '**************************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet id
        Call Buffer.ReadByte
        
        Dim RecordIndex As Byte

        Dim Obs         As String
        
        RecordIndex = Buffer.ReadByte
        Obs = Buffer.ReadASCIIString
        
        If Not (.flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster)) Then
            'Agregamos la observacion
            Call AddObs(UserIndex, RecordIndex, Obs)
            
            'Actualizamos la informacion
            Call WriteRecordDetails(UserIndex, RecordIndex)

        End If

        Call .incomingData.CopyBuffer(Buffer)

    End With
        
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "RecordRemove" message.
'
' @param UserIndex The index of the user sending the message.

Public Sub HandleRecordRemove(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Amraphen
    'Last Modification: 29/11/2010
    '
    '***************************************************
    Dim RecordIndex As Integer

    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
    
        RecordIndex = .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster) Then Exit Sub
        
        'Solo dioses pueden remover los seguimientos, los otros reciben una advertencia:
        If (.flags.Privilegios And PlayerType.Dios) Then
            Call RemoveRecord(RecordIndex)
            Call WriteShowMessageBox(UserIndex, "Se ha eliminado el seguimiento.")
            Call WriteRecordList(UserIndex)
        Else
            Call WriteShowMessageBox(UserIndex, "Solo los dioses pueden eliminar seguimientos.")

        End If

    End With

End Sub

''
' Handles the "RecordListRequest" message.
'
' @param UserIndex The index of the user sending the message.
            
Public Sub HandleRecordListRequest(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Amraphen
    'Last Modification: 29/11/2010
    '
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.RoleMaster) Then Exit Sub

        Call WriteRecordList(UserIndex)

    End With

End Sub


''
' Handles the "SetCharDescription" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSetCharDescription(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim tUser As Integer

        Dim Desc  As String
        
        Desc = Buffer.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin)) <> 0 Or (.flags.Privilegios And PlayerType.RoleMaster) <> 0 Then
            tUser = .flags.TargetUser

            If tUser > 0 Then
                UserList(tUser).DescRM = Desc
            Else
                Call WriteConsoleMsg(UserIndex, "Haz click sobre un personaje antes.", FontTypeNames.FONTTYPE_INFO)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "ForceMIDIToMap" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HanldeForceMIDIToMap(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim musicID As Byte

        Dim Mapa   As Integer
        
        musicID = .incomingData.ReadByte
        Mapa = .incomingData.ReadInteger
        
        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then

            'Si el mapa no fue enviado tomo el actual
            If Not InMapBounds(Mapa, 50, 50) Then
                Mapa = .Pos.Map

            End If
        
            If musicID = 0 Then
                'Ponemos el default del mapa
                Call SendData(SendTarget.toMap, Mapa, PrepareMessagePlayMusic(MapZonas(.Pos.Map).music, UserZonaId(UserIndex)))
            Else
                'Ponemos el pedido por el GM
                Call SendData(SendTarget.toMap, Mapa, PrepareMessagePlayMusic(musicID))

            End If

        End If

    End With

End Sub

''
' Handles the "ForceWAVEToMap" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleForceWAVEToMap(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 6 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim waveID As Byte

        Dim Mapa   As Integer

        Dim X      As Integer

        Dim Y      As Integer
        
        waveID = .incomingData.ReadByte()
        Mapa = .incomingData.ReadInteger()
        X = .incomingData.ReadInteger()
        Y = .incomingData.ReadInteger()
        
        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then

            'Si el mapa no fue enviado tomo el actual
            If Not InMapBounds(Mapa, X, Y) Then
                Mapa = .Pos.Map
                X = .Pos.X
                Y = .Pos.Y

            End If
            
            'Ponemos el pedido por el GM
            Call SendData(SendTarget.toMap, Mapa, PrepareMessagePlayWave(waveID, X, Y))

        End If

    End With

End Sub

''
' Handles the "RoyalArmyMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRoyalArmyMessage(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim Message As String

        Message = Buffer.ReadASCIIString()
        
        'Solo dioses, admins, semis y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then
            Call SendData(SendTarget.ToRealYRMs, 0, PrepareMessageConsoleMsg("EJERCITO REAL> " & Message, FontTypeNames.FONTTYPE_TALK))

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "ChaosLegionMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleChaosLegionMessage(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim Message As String

        Message = Buffer.ReadASCIIString()
        
        'Solo dioses, admins, semis y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then
            Call SendData(SendTarget.ToCaosYRMs, 0, PrepareMessageConsoleMsg("FUERZAS DEL CAOS> " & Message, FontTypeNames.FONTTYPE_TALK))

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "CitizenMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCitizenMessage(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim Message As String

        Message = Buffer.ReadASCIIString()
        
        'Solo dioses, admins, semis y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then
            Call SendData(SendTarget.ToCiudadanosYRMs, 0, PrepareMessageConsoleMsg("CIUDADANOS> " & Message, FontTypeNames.FONTTYPE_TALK))

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "CriminalMessage" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCriminalMessage(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim Message As String

        Message = Buffer.ReadASCIIString()
        
        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then
            Call SendData(SendTarget.ToCriminalesYRMs, 0, PrepareMessageConsoleMsg("CRIMINALES> " & Message, FontTypeNames.FONTTYPE_TALK))

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "TalkAsNPC" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTalkAsNPC(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/29/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim Message As String

        Message = Buffer.ReadASCIIString()
        
        'Solo dioses, admins y RMS
        If .flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin Or PlayerType.RoleMaster) Then

            'Asegurarse haya un NPC seleccionado
            If .flags.TargetNPC > 0 Then
                Call SendData(SendTarget.ToNPCArea, .flags.TargetNPC, PrepareMessageChatOverHead(Message, Npclist(.flags.TargetNPC).Char.CharIndex, 255, 255, 255))
            Else
                Call WriteConsoleMsg(UserIndex, "Debes seleccionar el NPC por el que quieres hablar antes de usar este comando.", FontTypeNames.FONTTYPE_INFO)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "DestroyAllItemsInArea" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDestroyAllItemsInArea(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Dim X       As Long

        Dim Y       As Long

        Dim bIsExit As Boolean
        
        For Y = .Pos.Y - MinYBorder + 1 To .Pos.Y + MinYBorder - 1
            For X = .Pos.X - MinXBorder + 1 To .Pos.X + MinXBorder - 1

                If X > 0 And Y > 0 And X < 101 And Y < 101 Then
                    If MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex > 0 Then

                        If ItemNoEsDeMapa(MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex) Then
                            Call EraseObj(MAX_INVENTORY_OBJS, .Pos.Map, X, Y)

                        End If

                    End If

                End If

            Next X
        Next Y
        
        Call LogGM(UserList(UserIndex).Name, "/MASSDEST")

    End With

End Sub

''
' Handles the "AcceptRoyalCouncilMember" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleAcceptRoyalCouncilMember(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim username As String

        Dim tUser    As Integer

        Dim LoopC    As Byte
        
        username = Buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            tUser = NameIndex(username)

            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline", FontTypeNames.FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.Toall, 0, PrepareMessageConsoleMsg(username & " fue aceptado en el honorable Consejo Real de Belleuve.", FontTypeNames.FONTTYPE_CONSEJO))

                With UserList(tUser)

                    If .flags.Privilegios And PlayerType.ChaosCouncil Then .flags.Privilegios = .flags.Privilegios - PlayerType.ChaosCouncil
                    If Not .flags.Privilegios And PlayerType.RoyalCouncil Then .flags.Privilegios = .flags.Privilegios + PlayerType.RoyalCouncil
                    
                    Call WarpUserChar(tUser, .Pos.Map, .Pos.X, .Pos.Y, False)

                End With

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "ChaosCouncilMember" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleAcceptChaosCouncilMember(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim username As String

        Dim tUser    As Integer

        Dim LoopC    As Byte
        
        username = Buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            tUser = NameIndex(username)

            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline", FontTypeNames.FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.Toall, 0, PrepareMessageConsoleMsg(username & " fue aceptado en el Concilio de las Sombras.", FontTypeNames.FONTTYPE_CONSEJO))
                
                With UserList(tUser)

                    If .flags.Privilegios And PlayerType.RoyalCouncil Then .flags.Privilegios = .flags.Privilegios - PlayerType.RoyalCouncil
                    If Not .flags.Privilegios And PlayerType.ChaosCouncil Then .flags.Privilegios = .flags.Privilegios + PlayerType.ChaosCouncil

                    Call WarpUserChar(tUser, .Pos.Map, .Pos.X, .Pos.Y, False)

                End With

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "ItemsInTheFloor" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleItemsInTheFloor(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 26/09/2024
    '26/09/2024 - Lorwik: Ahora solo recorre 50 a cada lado al rededor del que ejecuto el comando.
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Dim tObj  As Integer
        Dim lista As String
        
        Dim X     As Long
        Dim Y     As Long

        Dim MinX  As Integer
        Dim MaxX  As Integer
        Dim MinY  As Integer
        Dim MaxY  As Integer
        
        With UserList(UserIndex)
        
            MinX = .Pos.X - 50
            MaxX = .Pos.X + 50
            MinY = .Pos.Y - 50
            MaxY = .Pos.Y + 50
            
            If MinX < XMinMapSize Then MinX = XMinMapSize + 1
            If MaxX > XMaxMapSize Then MaxX = XMaxMapSize - 1
            If MinY < YMinMapSize Then MinY = YMinMapSize + 1
            If MaxY > YMaxMapSize Then MaxY = YMaxMapSize - 1
        
            For X = MinX To MaxX
                For Y = MinY To MaxY
                    tObj = MapData(.Pos.Map, X, Y).ObjInfo.ObjIndex
                
                    If tObj > 0 Then
                        If ObjData(tObj).OBJType <> eOBJType.otDestruible Then
                            Call WriteConsoleMsg(UserIndex, "(" & X & "," & Y & ") " & ObjData(tObj).Name, FontTypeNames.FONTTYPE_INFO)

                        End If

                    End If

                Next Y
            Next X
        
        End With

    End With

End Sub

''
' Handles the "MakeDumb" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleMakeDumb(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim username As String

        Dim tUser    As Integer
        
        username = Buffer.ReadASCIIString()
        
        If ((.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Or ((.flags.Privilegios And (PlayerType.SemiDios Or PlayerType.RoleMaster)) = (PlayerType.SemiDios Or PlayerType.RoleMaster))) Then
            tUser = NameIndex(username)

            'para deteccion de aoice
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteDumb(tUser)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "MakeDumbNoMore" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleMakeDumbNoMore(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim username As String

        Dim tUser    As Integer
        
        username = Buffer.ReadASCIIString()
        
        If ((.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Or ((.flags.Privilegios And (PlayerType.SemiDios Or PlayerType.RoleMaster)) = (PlayerType.SemiDios Or PlayerType.RoleMaster))) Then
            tUser = NameIndex(username)

            'para deteccion de aoice
            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteDumbNoMore(tUser)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "DumpIPTables" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDumpIPTables(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Call SecurityIp.DumpTables

    End With

End Sub

''
' Handles the "CouncilKick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCouncilKick(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim username As String

        Dim tUser    As Integer
        
        username = Buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            tUser = NameIndex(username)

            If tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_INFO)

            Else

                With UserList(tUser)

                    If .flags.Privilegios And PlayerType.RoyalCouncil Then
                        Call WriteConsoleMsg(tUser, "Has sido echado del consejo de Belleuve.", FontTypeNames.FONTTYPE_TALK)
                        .flags.Privilegios = .flags.Privilegios - PlayerType.RoyalCouncil
                        
                        Call WarpUserChar(tUser, .Pos.Map, .Pos.X, .Pos.Y, False)
                        Call SendData(SendTarget.Toall, 0, PrepareMessageConsoleMsg(username & " fue expulsado del consejo de Belleuve.", FontTypeNames.FONTTYPE_CONSEJO))

                    End If
                    
                    If .flags.Privilegios And PlayerType.ChaosCouncil Then
                        Call WriteConsoleMsg(tUser, "Has sido echado del Concilio de las Sombras.", FontTypeNames.FONTTYPE_TALK)
                        .flags.Privilegios = .flags.Privilegios - PlayerType.ChaosCouncil
                        
                        Call WarpUserChar(tUser, .Pos.Map, .Pos.X, .Pos.Y, False)
                        Call SendData(SendTarget.Toall, 0, PrepareMessageConsoleMsg(username & " fue expulsado del Concilio de las Sombras.", FontTypeNames.FONTTYPE_CONSEJO))

                    End If

                End With

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "SetTrigger" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleSetTrigger(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim tTrigger As Byte

        Dim tLog     As String
        
        tTrigger = .incomingData.ReadByte()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        If tTrigger >= 0 Then
            MapData(.Pos.Map, .Pos.X, .Pos.Y).Trigger = tTrigger
            tLog = "Trigger " & tTrigger & " en mapa " & .Pos.Map & " " & .Pos.X & "," & .Pos.Y
            
            Call LogGM(.Name, tLog)
            Call WriteConsoleMsg(UserIndex, tLog, FontTypeNames.FONTTYPE_INFO)

        End If

    End With

End Sub

''
' Handles the "AskTrigger" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleAskTrigger(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 04/13/07
    '
    '***************************************************
    Dim tTrigger As Byte
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        tTrigger = MapData(.Pos.Map, .Pos.X, .Pos.Y).Trigger
        
        Call LogGM(.Name, "Miro el trigger en " & .Pos.Map & "," & .Pos.X & "," & .Pos.Y & ". Era " & tTrigger)
        
        Call WriteConsoleMsg(UserIndex, "Trigger " & tTrigger & " en mapa " & .Pos.Map & " " & .Pos.X & ", " & .Pos.Y, FontTypeNames.FONTTYPE_INFO)

    End With

End Sub

''
' Handles the "BannedIPList" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBannedIPList(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Dim lista As String

        Dim LoopC As Long
        
        Call LogGM(.Name, "/BANIPLIST")
        
        For LoopC = 1 To BanIps.Count
            lista = lista & BanIps.Item(LoopC) & ", "
        Next LoopC
        
        If LenB(lista) <> 0 Then lista = Left$(lista, Len(lista) - 2)
        
        Call WriteConsoleMsg(UserIndex, lista, FontTypeNames.FONTTYPE_INFO)

    End With

End Sub

''
' Handles the "BannedIPReload" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBannedIPReload(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call BanIpGuardar
        Call BanIpCargar

    End With

End Sub

''
' Handles the "GuildBan" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleGuildBan(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim GuildName   As String

        Dim cantMembers As Integer

        Dim LoopC       As Long

        Dim member      As String

        Dim Count       As Byte

        Dim tIndex      As Integer

        Dim tFile       As String
        
        GuildName = Buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            tFile = App.Path & "\guilds\" & GuildName & "-members.mem"
            
            If Not FileExist(tFile) Then
                Call WriteConsoleMsg(UserIndex, "No existe el clan: " & GuildName, FontTypeNames.FONTTYPE_INFO)
            Else
                Call SendData(SendTarget.Toall, 0, PrepareMessageConsoleMsg(.Name & " baneo al clan " & UCase$(GuildName), FontTypeNames.FONTTYPE_GUILD))
                
                'baneamos a los miembros
                Call LogGM(.Name, "BANCLAN a " & UCase$(GuildName))
                
                cantMembers = val(GetVar(tFile, "INIT", "NroMembers"))
                
                For LoopC = 1 To cantMembers
                    member = GetVar(tFile, "Members", "Member" & LoopC)
                    'member es la victima
                    Call Ban(member, "Administracion del servidor", "Clan Banned")
                    
                    Call SendData(SendTarget.Toall, 0, PrepareMessageConsoleMsg("   " & member & "<" & GuildName & "> ha sido expulsado del servidor.", FontTypeNames.FONTTYPE_FIGHT))
                    
                    tIndex = NameIndex(member)

                    If tIndex > 0 Then
                        'esta online
                        UserList(tIndex).flags.Ban = 1
                        Call CloseUser(tIndex)

                    End If

                    Call SaveBan(member, "BAN AL CLAN: " & GuildName, LCase$(.Name))
                Next LoopC

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "BanIP" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleBanIP(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 07/02/09
    'Agregado un CopyBuffer porque se producia un bucle
    'inifito al intentar banear una ip ya baneada. (NicoNZ)
    '07/02/09 Pato - Ahora no es posible saber si un gm esta o no online.
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 6 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim bannedIP As String

        Dim tUser    As Integer

        Dim Reason   As String

        Dim i        As Long
        
        ' Is it by ip??
        If Buffer.ReadBoolean() Then
            bannedIP = Buffer.ReadByte() & "."
            bannedIP = bannedIP & Buffer.ReadByte() & "."
            bannedIP = bannedIP & Buffer.ReadByte() & "."
            bannedIP = bannedIP & Buffer.ReadByte()
        Else
            tUser = NameIndex(Buffer.ReadASCIIString())
            
            If tUser > 0 Then bannedIP = UserList(tUser).IP

        End If
        
        Reason = Buffer.ReadASCIIString()
        
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then
            If LenB(bannedIP) > 0 Then
                Call LogGM(.Name, "/BanIP " & bannedIP & " por " & Reason)
                
                If BanIpBuscar(bannedIP) > 0 Then
                    Call WriteConsoleMsg(UserIndex, "La IP " & bannedIP & " ya se encuentra en la lista de bans.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call BanIpAgrega(bannedIP)
                    Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.Name & " baneo la IP " & bannedIP & " por " & Reason, FontTypeNames.FONTTYPE_FIGHT))
                    
                    'Find every player with that ip and ban him!
                    For i = 1 To LastUser

                        If UserList(i).ConnIDValida Then
                            If UserList(i).IP = bannedIP Then
                                Call BanCharacter(UserIndex, UserList(i).Name, "IP POR " & Reason)

                            End If

                        End If

                    Next i

                End If

            ElseIf tUser <= 0 Then
                Call WriteConsoleMsg(UserIndex, "El personaje no esta online.", FontTypeNames.FONTTYPE_INFO)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "UnbanIP" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleUnbanIP(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 12/30/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim bannedIP As String
        
        bannedIP = .incomingData.ReadByte() & "."
        bannedIP = bannedIP & .incomingData.ReadByte() & "."
        bannedIP = bannedIP & .incomingData.ReadByte() & "."
        bannedIP = bannedIP & .incomingData.ReadByte()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        If BanIpQuita(bannedIP) Then
            Call WriteConsoleMsg(UserIndex, "La IP """ & bannedIP & """ se ha quitado de la lista de bans.", FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(UserIndex, "La IP """ & bannedIP & """ NO se encuentra en la lista de bans.", FontTypeNames.FONTTYPE_INFO)

        End If

    End With

End Sub

''
' Handles the "CreateItem" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleCreateItem(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 11/02/2011
    'maTih.- : Ahora se puede elegir, la cantidad a crear.
    '***************************************************
    
    On Error GoTo errHandler
    
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    With UserList(UserIndex)
        
        ' Recibo el ID del paquete
        Call .incomingData.ReadByte

        Dim tObj    As Integer: tObj = .incomingData.ReadInteger()
        Dim Cuantos As Integer: Cuantos = .incomingData.ReadInteger()
        
        ' Es Game-Master?
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        ' Si hace mas de 10000, lo sacamos cagando.
        If Cuantos > 10000 Then Call WriteConsoleMsg(UserIndex, "Estas tratando de crear demasiado, como mucho podes crear 10.000 unidades.", FontTypeNames.FONTTYPE_TALK): Exit Sub
        
        ' El indice proporcionado supera la cantidad minima o total de items existentes en el juego?
        If tObj < 1 Or tObj > NumObjDatas Then Exit Sub

        ' El nombre del objeto es nulo?
        If LenB(ObjData(tObj).Name) = 0 Then Exit Sub

        Dim Objeto As obj
        
        With Objeto
            .Amount = Cuantos
            .ObjIndex = tObj
        End With
        
        ' Chequeo si el objeto es AGARRABLE(para las puertas, arboles y demas objs. que no deberian estar en el inventario)
        '   0 = SI
        '   1 = NO
        If ObjData(tObj).Agarrable = 0 Then
            ' Trato de meterlo en el inventario.
            If MeterItemEnInventario(UserIndex, Objeto) Then
                Call WriteConsoleMsg(UserIndex, "Has creado " & Objeto.Amount & " unidades de " & ObjData(tObj).Name & ".", FontTypeNames.FONTTYPE_INFO)
            Else
                ' Si no hay espacio, lo tiro al piso.
                Call TirarItemAlPiso(.Pos, Objeto)
                Call WriteConsoleMsg(UserIndex, "No tenes espacio en tu inventario para crear el item.", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(UserIndex, "ATENCION: CREASTE [" & Cuantos & "] ITEMS, TIRE E INGRESE /DEST EN CONSOLA PARA DESTRUIR LOS QUE NO NECESITE!!", FontTypeNames.FONTTYPE_GUILD)
            End If
        Else
            ' Crear el item NO AGARRARBLE y tirarlo al piso.
            Call TirarItemAlPiso(.Pos, Objeto)
            Call WriteConsoleMsg(UserIndex, "ATENCION: CREASTE [" & Cuantos & "] ITEMS, TIRE E INGRESE /DEST EN CONSOLA PARA DESTRUIR LOS QUE NO NECESITE!!", FontTypeNames.FONTTYPE_GUILD)
        End If
        
        ' Lo registro en los logs.
        Call LogGM(.Name, "/CI: " & tObj & " - [Nombre del Objeto: " & ObjData(tObj).Name & "] - [Cantidad : " & Cuantos & "]")
        
    End With
    
errHandler:
    If Err.Number <> 0 Then
        Call LogError("Error en HandleCreateItem " & Err.Number & " " & Err.description)
    End If
End Sub

''
' Handles the "DestroyItems" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDestroyItems(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Dim Mapa As Integer

        Dim X    As Byte

        Dim Y    As Byte
        
        Mapa = .Pos.Map
        X = .Pos.X
        Y = .Pos.Y
        
        Dim ObjIndex As Integer

        ObjIndex = MapData(Mapa, X, Y).ObjInfo.ObjIndex
        
        If ObjIndex = 0 Then Exit Sub
        
        Call LogGM(.Name, "/DEST " & ObjIndex & " en mapa " & Mapa & " (" & X & "," & Y & "). Cantidad: " & MapData(Mapa, X, Y).ObjInfo.Amount)
        
        If ObjData(ObjIndex).OBJType = eOBJType.otTeleport And MapData(Mapa, X, Y).TileExit.Map > 0 Then
            
            Call WriteConsoleMsg(UserIndex, "No puede destruir teleports asi. Utilice /DT.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        Call EraseObj(10000, Mapa, X, Y)

    End With

End Sub

''
' Handles the "ChaosLegionKick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleChaosLegionKick(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim username As String

        Dim tUser    As Integer
        
        username = Buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Or .flags.PrivEspecial Then
            
            If (InStrB(username, "\") <> 0) Then
                username = Replace(username, "\", "")

            End If

            If (InStrB(username, "/") <> 0) Then
                username = Replace(username, "/", "")

            End If

            tUser = NameIndex(username)
            
            Call LogGM(.Name, "ECHO DEL CAOS A: " & username)
    
            If tUser > 0 Then
                Call ExpulsarFaccionCaos(tUser, True)
                UserList(tUser).Faccion.Reenlistadas = 200
                Call WriteConsoleMsg(UserIndex, username & " expulsado de las fuerzas del caos y prohibida la reenlistada.", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(tUser, .Name & " te ha expulsado en forma definitiva de las fuerzas del caos.", FontTypeNames.FONTTYPE_FIGHT)
            Else

                If PersonajeExiste(username) Then
                    Call KickUserChaosLegion(username)
                    Call WriteConsoleMsg(UserIndex, username & " expulsado de las fuerzas del caos y prohibida la reenlistada.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, username & " inexistente.", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "RoyalArmyKick" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRoyalArmyKick(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim username As String

        Dim tUser    As Integer
        
        username = Buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Or .flags.PrivEspecial Then
            
            If (InStrB(username, "\") <> 0) Then
                username = Replace(username, "\", "")

            End If

            If (InStrB(username, "/") <> 0) Then
                username = Replace(username, "/", "")

            End If

            tUser = NameIndex(username)
            
            Call LogGM(.Name, "ECHO DE LA REAL A: " & username)
            
            If tUser > 0 Then
                Call ExpulsarFaccionReal(tUser, True)
                UserList(tUser).Faccion.Reenlistadas = 200
                Call WriteConsoleMsg(UserIndex, username & " expulsado de las fuerzas reales y prohibida la reenlistada.", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(tUser, .Name & " te ha expulsado en forma definitiva de las fuerzas reales.", FontTypeNames.FONTTYPE_FIGHT)
            Else

                If PersonajeExiste(username) Then
                    Call KickUserRoyalArmy(username)
                    Call WriteConsoleMsg(UserIndex, username & " expulsado de las fuerzas reales y prohibida la reenlistada.", FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, username & " inexistente.", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "ForceMIDIAll" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleForceMUSICAll(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim musicID As Byte

        musicID = .incomingData.ReadByte()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Call SendData(SendTarget.Toall, 0, PrepareMessageConsoleMsg(.Name & " broadcast musica MUSIC: " & musicID, FontTypeNames.FONTTYPE_SERVER))
        
        Call SendData(SendTarget.Toall, 0, PrepareMessagePlayMusic(musicID))

    End With

End Sub

''
' Handles the "ForceWAVEAll" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleForceWAVEAll(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte

        Dim waveID As Byte

        waveID = .incomingData.ReadByte()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Call SendData(SendTarget.Toall, 0, PrepareMessagePlayWave(waveID, NO_3D_SOUND, NO_3D_SOUND))

    End With

End Sub

''
' Handles the "RemovePunishment" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleRemovePunishment(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 1/05/07
    'Pablo (ToxicWaste): 1/05/07, You can now edit the punishment.
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 6 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim username   As String

        Dim punishment As Byte

        Dim NewText    As String
        
        username = Buffer.ReadASCIIString()
        punishment = Buffer.ReadByte
        NewText = Buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            If LenB(username) = 0 Then
                Call WriteConsoleMsg(UserIndex, "Utilice /borrarpena Nick@NumeroDePena@NuevaPena", FontTypeNames.FONTTYPE_INFO)
            Else

                If (InStrB(username, "\") <> 0) Then
                    username = Replace(username, "\", "")

                End If

                If (InStrB(username, "/") <> 0) Then
                    username = Replace(username, "/", "")

                End If
                
                If PersonajeExiste(username) Then
                    Call LogGM(.Name, " borro la pena: " & punishment & " de " & username & " y la cambio por: " & NewText)

                    Call AlterUserPunishment(username, punishment, LCase$(.Name) & ": <" & NewText & "> " & Date & " " & time)
                    Call WriteConsoleMsg(UserIndex, "Pena modificada.", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "TileBlockedToggle" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleTileBlockedToggle(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub

        Call LogGM(.Name, "/BLOQ")
        
        If MapData(.Pos.Map, .Pos.X, .Pos.Y).Blocked = 0 Then
            MapData(.Pos.Map, .Pos.X, .Pos.Y).Blocked = 1
        Else
            MapData(.Pos.Map, .Pos.X, .Pos.Y).Blocked = 0

        End If
        
        Call Bloquear(True, .Pos.Map, .Pos.X, .Pos.Y, MapData(.Pos.Map, .Pos.X, .Pos.Y).Blocked)

    End With

End Sub

''
' Handles the "KillNPCNoRespawn" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleKillNPCNoRespawn(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        If .flags.TargetNPC = 0 Then Exit Sub
        
        Call QuitarNPC(.flags.TargetNPC)
        Call LogGM(.Name, "/MATA " & Npclist(.flags.TargetNPC).Name)

    End With

End Sub

''
' Handles the "KillAllNearbyNPCs" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleKillAllNearbyNPCs(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Dim X As Long

        Dim Y As Long
        
        For Y = .Pos.Y - MinYBorder + 1 To .Pos.Y + MinYBorder - 1
            For X = .Pos.X - MinXBorder + 1 To .Pos.X + MinXBorder - 1

                If X > 0 And Y > 0 And X < 101 And Y < 101 Then
                    If MapData(.Pos.Map, X, Y).NPCIndex > 0 Then Call QuitarNPC(MapData(.Pos.Map, X, Y).NPCIndex)

                End If

            Next X
        Next Y

        Call LogGM(.Name, "/MASSKILL")

    End With

End Sub

''
' Handles the "LastIP" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleLastIP(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 12/30/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim username   As String

        Dim lista      As String

        Dim LoopC      As Byte

        Dim priv       As Integer

        Dim validCheck As Boolean
        
        priv = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero
        username = Buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then

            'Handle special chars
            If (InStrB(username, "\") <> 0) Then
                username = Replace(username, "\", "")

            End If

            If (InStrB(username, "\") <> 0) Then
                username = Replace(username, "/", "")

            End If

            If (InStrB(username, "+") <> 0) Then
                username = Replace(username, "+", " ")

            End If
            
            'Only Gods and Admins can see the ips of adminsitrative characters. All others can be seen by every adminsitrative char.
            If NameIndex(username) > 0 Then
                validCheck = (UserList(NameIndex(username)).flags.Privilegios And priv) = 0 Or (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0
            Else
                validCheck = (UserDarPrivilegioLevel(username) And priv) = 0 Or (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0

            End If
            
            If validCheck Then
                Call LogGM(.Name, "/LASTIP " & username)
                
                If PersonajeExiste(username) Then
                    lista = "Las ultimas IPs con las que " & username & " se conecto son:" & vbCrLf & GetUserLastIps(username)
                    Call WriteConsoleMsg(UserIndex, lista, FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, "Charfile """ & username & """ inexistente.", FontTypeNames.FONTTYPE_INFO)

                End If

            Else
                Call WriteConsoleMsg(UserIndex, username & " es de mayor jerarquia que vos.", FontTypeNames.FONTTYPE_INFO)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "ChatColor" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleChatColor(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/23/06
    'Last modified by: Juan Martin Sotuyo Dodero (Maraxus)
    'Change the user`s chat color
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Dim r As Byte
        Dim g As Byte
        Dim b As Byte
        
        r = .incomingData.ReadByte()
        g = .incomingData.ReadByte()
        b = .incomingData.ReadByte()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.RoleMaster)) Then
            .flags.ChatColor.r = r
            .flags.ChatColor.g = g
            .flags.ChatColor.b = b
        End If

    End With

End Sub

''
' Handles the "Ignored" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleIgnored(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/23/06
    'Ignore the user
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero) Then
            .flags.AdminPerseguible = Not .flags.AdminPerseguible

        End If

    End With

End Sub

''
' Handles the "CheckSlot" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleCheckSlot(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Pablo (ToxicWaste)
    'Last Modification: 07/06/2010
    'Check one Users Slot in Particular from Inventory
    '07/06/2010: ZaMa - Ahora no se puede usar para saber si hay dioses/admins online.
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        'Reads the UserName and Slot Packets
        Dim username         As String

        Dim Slot             As Byte

        Dim tIndex           As Integer
        
        Dim UserIsAdmin      As Boolean

        Dim OtherUserIsAdmin As Boolean
                
        username = Buffer.ReadASCIIString() 'Que UserName?
        Slot = Buffer.ReadByte() 'Que Slot?
        
        UserIsAdmin = (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0

        If (.flags.Privilegios And PlayerType.SemiDios) <> 0 Or UserIsAdmin Then
            
            Call LogGM(.Name, .Name & " Checkeo el slot " & Slot & " de " & username)
            
            tIndex = NameIndex(username)  'Que user index?
            OtherUserIsAdmin = EsDios(username) Or EsAdmin(username)
            
            If tIndex > 0 Then
                If UserIsAdmin Or Not OtherUserIsAdmin Then
                    If Slot > 0 And Slot <= UserList(tIndex).CurrentInventorySlots Then
                        If UserList(tIndex).Invent.Object(Slot).ObjIndex > 0 Then
                            Call WriteConsoleMsg(UserIndex, " Objeto " & Slot & ") " & ObjData(UserList(tIndex).Invent.Object(Slot).ObjIndex).Name & " Cantidad:" & UserList(tIndex).Invent.Object(Slot).Amount, FontTypeNames.FONTTYPE_INFO)
                        Else
                            Call WriteConsoleMsg(UserIndex, "No hay ningUn objeto en slot seleccionado.", FontTypeNames.FONTTYPE_INFO)

                        End If

                    Else
                        Call WriteConsoleMsg(UserIndex, "Slot Invalido.", FontTypeNames.FONTTYPE_TALK)

                    End If

                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes ver slots de un dios o admin.", FontTypeNames.FONTTYPE_INFO)

                End If

            Else

                If UserIsAdmin Or Not OtherUserIsAdmin Then
                    Call WriteConsoleMsg(UserIndex, "Usuario offline.", FontTypeNames.FONTTYPE_TALK)
                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes ver slots de un dios o admin.", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the "ResetAutoUpdate" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleResetAutoUpdate(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/23/06
    'Reset the AutoUpdate
    '***************************************************
    With UserList(UserIndex)
        'Remove packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        Call WriteConsoleMsg(UserIndex, "TID: " & CStr(ReiniciarAutoUpdate()), FontTypeNames.FONTTYPE_INFO)

    End With

End Sub

''
' Handles the "Restart" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleRestart(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/23/06
    'Restart the game
    '***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
    
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        'time and Time BUG!
        Call LogGM(.Name, .Name & " reinicio el mundo.")
        
        Call ReiniciarServidor(True)

    End With

End Sub

''
' Handles the "ReloadObjects" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleReloadObjects(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/23/06
    'Reload the objects
    '***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha recargado los objetos.")
        
        Call LoadOBJData

    End With

End Sub

''
' Handles the "ReloadSpells" message.
'
' @param    userIndex The index of the user sending the message.

Public Sub HandleReloadSpells(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/23/06
    'Reload the spells
    '***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha recargado los hechizos.")
        
        Call CargarHechizos

    End With

End Sub

''
' Handle the "ReloadServerIni" message.
'
' @param userIndex The index of the user sending the message

Public Sub HandleReloadServerIni(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/23/06
    'Reload the Server`s INI
    '***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha recargado los INITs.")
        
        Call LoadSini
        
        Call WriteConsoleMsg(UserIndex, "Server.ini actualizado correctamente", FontTypeNames.FONTTYPE_INFO)

    End With

End Sub

''
' Handle the "ReloadNPCs" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleReloadNPCs(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/23/06
    'Reload the Server`s NPC
    '***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
         
        Call LogGM(.Name, .Name & " ha recargado los NPCs.")
    
        Call CargaNpcsDat
    
        Call WriteConsoleMsg(UserIndex, "Npcs.dat recargado.", FontTypeNames.FONTTYPE_INFO)

    End With

End Sub

''
' Handle the "KickAllChars" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleKickAllChars(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/23/06
    'Kick all the chars that are online
    '***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha echado a todos los personajes.")
        
        Call EcharPjsNoPrivilegiados

    End With

End Sub

''
' Handle the "Night" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleNight(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/23/06
    'Last modified by: Juan Martin Sotuyo Dodero (Maraxus)
    '
    '***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios) Then Exit Sub
        
        DeNoche = Not DeNoche
        
        Dim i As Long
        
        For i = 1 To NumUsers

            If UserList(i).flags.UserLogged And UserList(i).ConnID > -1 Then
                Call EnviarNoche(i)

            End If

        Next i

    End With

End Sub

''
' Handle the "ShowServerForm" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleShowServerForm(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/23/06
    'Show the server form
    '***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha solicitado mostrar el formulario del servidor.")
        Call frmMain.mnuMostrar_Click

    End With

End Sub

''
' Handle the "CleanSOS" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleCleanSOS(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/23/06
    'Clean the SOS
    '***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha borrado los SOS.")
        
        Call Ayuda.Reset

    End With

End Sub

''
' Handle the "SaveChars" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleSaveChars(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/23/06
    'Save the characters
    '***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha guardado todos los chars.")
        
        Call mdParty.ActualizaExperiencias
        Call GuardarUsuarios

    End With

End Sub

''
' Handle the "ChangeMapZonasBackup" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeZonaBackup(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/24/06
    'Last modified by: Juan Martin Sotuyo Dodero (Maraxus)
    'Change the backup`s info of the map
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        Dim doTheBackUp As Boolean
        
        doTheBackUp = .incomingData.ReadBoolean()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) = 0 Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha cambiado la informacion sobre el BackUp.")
        
        'Change the boolean to byte in a fast way
        If doTheBackUp Then
            MapZonas(.Pos.Map, UserZonaId(UserIndex)).BackUp = 1
        Else
            MapZonas(.Pos.Map, UserZonaId(UserIndex)).BackUp = 0

        End If
        
        'Change the boolean to string in a fast way
        Call WriteVar(MapPath & "mapa" & .Pos.Map & ".dat", "Mapa" & .Pos.Map, "backup", MapZonas(.Pos.Map, UserZonaId(UserIndex)).BackUp)
        
        Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " Backup: " & MapZonas(.Pos.Map, UserZonaId(UserIndex)).BackUp, FontTypeNames.FONTTYPE_INFO)

    End With

End Sub

''
' Handle the "ChangeZonaPK" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeZonaPK(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/24/06
    'Last modified by: Juan Martin Sotuyo Dodero (Maraxus)
    'Change the pk`s info of the  map
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        Dim isMapPk As Boolean
        
        isMapPk = .incomingData.ReadBoolean()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) = 0 Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha cambiado la informacion sobre si es PK el mapa.")
        
        MapZonas(.Pos.Map, UserZonaId(UserIndex)).Pk = isMapPk
        
        'Change the boolean to string in a fast way
        Call WriteVar(MapPath & "mapa" & .Pos.Map & ".dat", "Mapa" & .Pos.Map, "Pk", IIf(isMapPk, "1", "0"))

        Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " PK: " & MapZonas(.Pos.Map, UserZonaId(UserIndex)).Pk, FontTypeNames.FONTTYPE_INFO)

    End With

End Sub

''
' Handle the "ChangeZonaRestricted" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeZonaRestricted(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Pablo (ToxicWaste)
    'Last Modification: 26/01/2007
    'Restringido -> Options: "NEWBIE", "NO", "ARMADA", "CAOS", "FACCION".
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    Dim tStr As String
    
    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove Packet ID
        Call Buffer.ReadByte
        
        tStr = Buffer.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            If tStr = "NEWBIE" Or tStr = "NO" Or tStr = "ARMADA" Or tStr = "CAOS" Or tStr = "FACCION" Then
                Call LogGM(.Name, .Name & " ha cambiado la informacion sobre si es restringido el mapa.")
                
                MapZonas(UserList(UserIndex).Pos.Map, UserZonaId(UserIndex)).Restringir = RestrictStringToByte(tStr)
                
                Call WriteVar(MapPath & "mapa" & UserList(UserIndex).Pos.Map & ".dat", "Mapa" & UserList(UserIndex).Pos.Map, "Restringir", tStr)
                Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " Restringido: " & RestrictByteToString(MapZonas(.Pos.Map, UserZonaId(UserIndex)).Restringir), FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "Opciones para restringir: 'NEWBIE', 'NO', 'ARMADA', 'CAOS', 'FACCION'", FontTypeNames.FONTTYPE_INFO)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handle the "ChangeZonaNoMagic" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeZonaNoMagic(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Pablo (ToxicWaste)
    'Last Modification: 26/01/2007
    'MagiaSinEfecto -> Options: "1" , "0".
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    Dim nomagic As Boolean
    
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        nomagic = .incomingData.ReadBoolean
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            Call LogGM(.Name, .Name & " ha cambiado la informacion sobre si esta permitido usar la magia el mapa.")
            MapZonas(UserList(UserIndex).Pos.Map, UserZonaId(UserIndex)).MagiaSinEfecto = nomagic
            Call WriteVar(MapPath & "mapa" & UserList(UserIndex).Pos.Map & ".dat", "Mapa" & UserList(UserIndex).Pos.Map, "MagiaSinEfecto", nomagic)
            Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " MagiaSinEfecto: " & MapZonas(.Pos.Map, UserZonaId(UserIndex)).MagiaSinEfecto, FontTypeNames.FONTTYPE_INFO)

        End If

    End With

End Sub

''
' Handle the "ChangeZonaNoInvi" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeZonaNoInvi(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Pablo (ToxicWaste)
    'Last Modification: 26/01/2007
    'InviSinEfecto -> Options: "1", "0"
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    Dim noinvi As Boolean
    
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        noinvi = .incomingData.ReadBoolean()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            Call LogGM(.Name, .Name & " ha cambiado la informacion sobre si esta permitido usar la invisibilidad en el mapa.")
            MapZonas(UserList(UserIndex).Pos.Map, UserZonaId(UserIndex)).InviSinEfecto = noinvi
            Call WriteVar(MapPath & "mapa" & UserList(UserIndex).Pos.Map & ".dat", "Mapa" & UserList(UserIndex).Pos.Map, "InviSinEfecto", noinvi)
            Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " InviSinEfecto: " & MapZonas(.Pos.Map, UserZonaId(UserIndex)).InviSinEfecto, FontTypeNames.FONTTYPE_INFO)

        End If

    End With

End Sub
            
''
' Handle the "ChangeZonaNoResu" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeZonaNoResu(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Pablo (ToxicWaste)
    'Last Modification: 26/01/2007
    'ResuSinEfecto -> Options: "1", "0"
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    Dim noresu As Boolean
    
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        noresu = .incomingData.ReadBoolean()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            Call LogGM(.Name, .Name & " ha cambiado la informacion sobre si esta permitido usar el resucitar en el mapa.")
            MapZonas(UserList(UserIndex).Pos.Map, UserZonaId(UserIndex)).ResuSinEfecto = noresu
            Call WriteVar(MapPath & "mapa" & UserList(UserIndex).Pos.Map & ".dat", "Mapa" & UserList(UserIndex).Pos.Map, "ResuSinEfecto", noresu)
            Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " ResuSinEfecto: " & MapZonas(.Pos.Map, UserZonaId(UserIndex)).ResuSinEfecto, FontTypeNames.FONTTYPE_INFO)

        End If

    End With

End Sub

''
' Handle the "ChangeZonaLand" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeZonaLand(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Pablo (ToxicWaste)
    'Last Modification: 26/01/2007
    'Terreno -> Opciones: "BOSQUE", "NIEVE", "DESIERTO", "CIUDAD", "CAMPO", "DUNGEON".
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    Dim tStr As String
    
    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove Packet ID
        Call Buffer.ReadByte
        
        tStr = Buffer.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            If tStr = "BOSQUE" Or tStr = "NIEVE" Or tStr = "DESIERTO" Or tStr = "CIUDAD" Or tStr = "CAMPO" Or tStr = "DUNGEON" Then
                Call LogGM(.Name, .Name & " ha cambiado la informacion del terreno del mapa.")
                
                MapZonas(UserList(UserIndex).Pos.Map, UserZonaId(UserIndex)).Terreno = TerrainStringToByte(tStr)
                
                Call WriteVar(MapPath & "mapa" & UserList(UserIndex).Pos.Map & ".dat", "Mapa" & UserList(UserIndex).Pos.Map, "Terreno", tStr)
                Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " Terreno: " & TerrainByteToString(MapZonas(.Pos.Map, UserZonaId(UserIndex)).Terreno), FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "Opciones para terreno: 'BOSQUE', 'NIEVE', 'DESIERTO', 'CIUDAD', 'CAMPO', 'DUNGEON'", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(UserIndex, "Igualmente, el Unico Util es 'NIEVE' ya que al ingresarlo, la gente muere de frio en el mapa.", FontTypeNames.FONTTYPE_INFO)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handle the "ChangeZonaZone" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeZonaZone(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Pablo (ToxicWaste)
    'Last Modification: 26/01/2007
    'Zona -> Opciones: "BOSQUE", "NIEVE", "DESIERTO", "CIUDAD", "CAMPO", "DUNGEON".
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    Dim tStr As String
    
    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove Packet ID
        Call Buffer.ReadByte
        
        tStr = Buffer.ReadASCIIString()
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            If tStr = "BOSQUE" Or tStr = "NIEVE" Or tStr = "DESIERTO" Or tStr = "CIUDAD" Or tStr = "CAMPO" Or tStr = "DUNGEON" Then
                Call LogGM(.Name, .Name & " ha cambiado la informacion de la zona del mapa.")
                MapZonas(UserList(UserIndex).Pos.Map, UserZonaId(UserIndex)).Zona = tStr
                Call WriteVar(MapPath & "mapa" & UserList(UserIndex).Pos.Map & ".dat", "Mapa" & UserList(UserIndex).Pos.Map, "Zona", tStr)
                Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " Zona: " & MapZonas(.Pos.Map, UserZonaId(UserIndex)).Zona, FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "Opciones para terreno: 'BOSQUE', 'NIEVE', 'DESIERTO', 'CIUDAD', 'CAMPO', 'DUNGEON'", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(UserIndex, "Igualmente, el Unico Util es 'DUNGEON' ya que al ingresarlo, NO se sentira el efecto de la lluvia en este mapa.", FontTypeNames.FONTTYPE_INFO)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub
            
''
' Handle the "ChangeZonaStealNp" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeZonaStealNpc(ByVal UserIndex As Integer)

    '***************************************************
    'Author: ZaMa
    'Last Modification: 25/07/2010
    'RoboNpcsPermitido -> Options: "1", "0"
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    Dim RoboNpc As Byte
    
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        RoboNpc = val(IIf(.incomingData.ReadBoolean(), 1, 0))
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            Call LogGM(.Name, .Name & " ha cambiado la informacion sobre si esta permitido robar npcs en el mapa.")
            
            MapZonas(UserList(UserIndex).Pos.Map, UserZonaId(UserIndex)).RoboNpcsPermitido = RoboNpc
            
            Call WriteVar(MapPath & "mapa" & UserList(UserIndex).Pos.Map & ".dat", "Mapa" & UserList(UserIndex).Pos.Map, "RoboNpcsPermitido", RoboNpc)
            Call WriteConsoleMsg(UserIndex, "Mapa " & .Pos.Map & " RoboNpcsPermitido: " & MapZonas(.Pos.Map, UserZonaId(UserIndex)).RoboNpcsPermitido, FontTypeNames.FONTTYPE_INFO)

        End If

    End With

End Sub
            
''
' Handle the "ChangeZonaNoOcultar" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeZonaNoOcultar(ByVal UserIndex As Integer)

    '***************************************************
    'Author: ZaMa
    'Last Modification: 18/09/2010
    'OcultarSinEfecto -> Options: "1", "0"
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    Dim NoOcultar As Byte

    Dim Mapa      As Integer
    
    With UserList(UserIndex)
    
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        NoOcultar = val(IIf(.incomingData.ReadBoolean(), 1, 0))
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            
            Mapa = .Pos.Map
            
            Call LogGM(.Name, .Name & " ha cambiado la informacion sobre si esta permitido ocultarse en el mapa " & Mapa & ".")
            
            MapZonas(Mapa, UserZonaId(UserIndex)).OcultarSinEfecto = NoOcultar

            Call WriteVar(MapPath & "mapa" & Mapa & ".dat", "Mapa" & Mapa, "OcultarSinEfecto", NoOcultar)
            Call WriteConsoleMsg(UserIndex, "Mapa " & Mapa & " OcultarSinEfecto: " & NoOcultar, FontTypeNames.FONTTYPE_INFO)

        End If
        
    End With
    
End Sub
           
''
' Handle the "ChangeZonaNoInvocar" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeZonaNoInvocar(ByVal UserIndex As Integer)

    '***************************************************
    'Author: ZaMa
    'Last Modification: 18/09/2010
    'InvocarSinEfecto -> Options: "1", "0"
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 2 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    Dim NoInvocar As Byte

    Dim Mapa      As Integer
    
    With UserList(UserIndex)
    
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        NoInvocar = val(IIf(.incomingData.ReadBoolean(), 1, 0))
        
        If (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) <> 0 Then
            
            Mapa = .Pos.Map
            
            Call LogGM(.Name, .Name & " ha cambiado la informacion sobre si esta permitido invocar en el mapa " & Mapa & ".")
            
            MapZonas(Mapa, UserZonaId(UserIndex)).InvocarSinEfecto = NoInvocar

            Call WriteVar(MapPath & "mapa" & Mapa & ".dat", "Mapa" & Mapa, "InvocarSinEfecto", NoInvocar)
            Call WriteConsoleMsg(UserIndex, "Mapa " & Mapa & " InvocarSinEfecto: " & NoInvocar, FontTypeNames.FONTTYPE_INFO)

        End If
        
    End With
    
End Sub

''
' Handle the "SaveMap" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleSaveMap(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/24/06
    'Saves the map
    '***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha guardado el mapa " & CStr(.Pos.Map))
        
        Call GrabarMapa(.Pos.Map, App.Path & "\WorldBackUp\Mapa" & .Pos.Map)
        
        Call WriteConsoleMsg(UserIndex, "Mapa Guardado.", FontTypeNames.FONTTYPE_INFO)

    End With

End Sub

''
' Handle the "ShowGuildMessages" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleShowGuildMessages(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/24/06
    'Last modified by: Juan Martin Sotuyo Dodero (Maraxus)
    'Allows admins to read guild messages
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim Guild As String
        
        Guild = Buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call modGuilds.GMEscuchaClan(UserIndex, Guild)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handle the "DoBackUp" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleDoBackUp(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/24/06
    'Show guilds messages
    '***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, .Name & " ha hecho un backup.")
        
        Call ES.DoBackUp 'Sino lo confunde con la id del paquete

    End With

End Sub

''
' Handle the "AlterName" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleAlterName(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 12/26/06
    'Change user name
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        'Reads the userName and newUser Packets
        Dim username     As String

        Dim newName      As String

        Dim changeNameUI As Integer

        Dim GuildIndex   As Integer
        
        username = Buffer.ReadASCIIString()
        newName = Buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Or .flags.PrivEspecial Then
            If LenB(username) = 0 Or LenB(newName) = 0 Then
                Call WriteConsoleMsg(UserIndex, "Usar: /ANAME origen@destino", FontTypeNames.FONTTYPE_INFO)
            Else
                changeNameUI = NameIndex(username)
                
                If changeNameUI > 0 Then
                    Call WriteConsoleMsg(UserIndex, "El Pj esta online, debe salir para hacer el cambio.", FontTypeNames.FONTTYPE_WARNING)
                Else

                    If Not PersonajeExiste(username) Then
                        Call WriteConsoleMsg(UserIndex, "El pj " & username & " es inexistente.", FontTypeNames.FONTTYPE_INFO)
                    Else

                        If GetUserGuildIndex(username) > 0 Then
                            Call WriteConsoleMsg(UserIndex, "El pj " & username & " pertenece a un clan, debe salir del mismo con /salirclan para ser transferido.", FontTypeNames.FONTTYPE_INFO)
                        Else

                            If Not PersonajeExiste(newName) Then
                                Call CopyUser(username, newName)

                                Call WriteConsoleMsg(UserIndex, "Transferencia exitosa.", FontTypeNames.FONTTYPE_INFO)
                                Call LogGM(.Name, "Ha cambiado de nombre al usuario " & username & ". Ahora se llama " & newName)
                            Else
                                Call WriteConsoleMsg(UserIndex, "El nick solicitado ya existe.", FontTypeNames.FONTTYPE_INFO)

                            End If

                        End If

                    End If

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handle the "HandleCreateNPC" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleCreateNPC(ByVal UserIndex As Integer)

    '**********************************************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 11/05/2019
    '26/09/2010: ZaMa - Ya no se pueden crear npcs pretorianos.
    '11/05/2019: Jopi - Se arreglo la comprobacion de NPC's pretorianos.
    '11/05/2019: Jopi - Se combino HandleCreateNPCWithRespawn() con este procedimiento.
    '**********************************************************************************
    If UserList(UserIndex).incomingData.Length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        Dim NPCIndex As Integer: NPCIndex = .incomingData.ReadInteger()
        Dim Respawn As Boolean: Respawn = .incomingData.ReadBoolean()
        
        'Nos fijamos que sea GM.
        If Not EsGm(UserIndex) Then Exit Sub
        
        'Nos fijamos si es pretoriano.
        If PRETORIANOS_ACTIVADO Then
            If Npclist(NPCIndex).NPCtype = eNPCType.Pretoriano Then
                Call WriteConsoleMsg(UserIndex, "No puedes sumonear miembros del clan pretoriano de esta forma, utiliza /CREARPRETORIANOS MAPA X Y.", FontTypeNames.FONTTYPE_WARNING)
                Exit Sub
    
            End If
        End If
        
        'Invocamos el NPC.
        If NPCIndex <> 0 Then
        
            NPCIndex = SpawnNpc(NPCIndex, .Pos, True, Respawn)
        
            Call LogGM(.Name, "Invoco " & IIf(Respawn, "con respawn", vbNullString) & " a " & Npclist(NPCIndex).Name & " [Indice: " & NPCIndex & "] en el mapa " & .Pos.Map)

        End If

    End With

End Sub

''
' Handle the "ImperialArmour" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleImperialArmour(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 12/24/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        Dim Index    As Byte

        Dim ObjIndex As Integer
        
        Index = .incomingData.ReadByte()
        ObjIndex = .incomingData.ReadInteger()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Select Case Index

            Case 1
                ArmaduraImperial1 = ObjIndex
            
            Case 2
                ArmaduraImperial2 = ObjIndex
            
            Case 3
                ArmaduraImperial3 = ObjIndex
            
            Case 4
                TunicaMagoImperial = ObjIndex

        End Select

    End With

End Sub

''
' Handle the "ChaosArmour" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChaosArmour(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 12/24/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        Dim Index    As Byte

        Dim ObjIndex As Integer
        
        Index = .incomingData.ReadByte()
        ObjIndex = .incomingData.ReadInteger()
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        Select Case Index

            Case 1
                ArmaduraCaos1 = ObjIndex
            
            Case 2
                ArmaduraCaos2 = ObjIndex
            
            Case 3
                ArmaduraCaos3 = ObjIndex
            
            Case 4
                TunicaMagoCaos = ObjIndex

        End Select

    End With

End Sub

''
' Handle the "NavigateToggle" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleNavigateToggle(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 01/12/07
    '
    '***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then Exit Sub
        
        If .flags.Navegando = 1 Then
            .flags.Navegando = 0
        Else
            .flags.Navegando = 1

        End If
        
        'Tell the client that we are navigating.
        Call WriteNavigateToggle(UserIndex)

    End With

End Sub

''
' Handle the "ServerOpenToUsersToggle" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleServerOpenToUsersToggle(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 12/24/06
    '
    '***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.RoleMaster) Then Exit Sub
        
        If ServerSoloGMs > 0 Then
            Call WriteConsoleMsg(UserIndex, "Servidor habilitado para todos.", FontTypeNames.FONTTYPE_INFO)
            ServerSoloGMs = 0
            frmMain.chkServerHabilitado.Value = vbUnchecked
        Else
            Call WriteConsoleMsg(UserIndex, "Servidor restringido a administradores.", FontTypeNames.FONTTYPE_INFO)
            ServerSoloGMs = 1
            frmMain.chkServerHabilitado.Value = vbChecked

        End If

    End With

End Sub

''
' Handle the "TurnOffServer" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleTurnOffServer(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 12/24/06
    'Turns off the server
    '***************************************************
    Dim handle As Integer
    
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios Or PlayerType.Dios Or PlayerType.RoleMaster) Then Exit Sub
        
        Call LogGM(.Name, "/APAGAR")
        Call SendData(SendTarget.Toall, 0, PrepareMessageConsoleMsg("" & .Name & " VA A APAGAR EL SERVIDOR!!!", FontTypeNames.FONTTYPE_FIGHT))
        
        'Log
        handle = FreeFile
        Open App.Path & "\logs\Main.log" For Append Shared As #handle
        
        Print #handle, Date & " " & time & " server apagado por " & .Name & ". "
        
        Close #handle
        
        Call CloseServer

    End With

End Sub

''
' Handle the "TurnCriminal" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleTurnCriminal(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 12/26/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim username As String

        Dim tUser    As Integer
        
        username = Buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call LogGM(.Name, "/CONDEN " & username)
            
            tUser = NameIndex(username)

            If tUser > 0 Then Call VolverCriminal(tUser)

        End If
                
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handle the "ResetFactions" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleResetFactions(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 06/09/09
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim username As String

        Dim tUser    As Integer

        Dim Char     As String
        
        username = Buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Or .flags.PrivEspecial Then
            Call LogGM(.Name, "/RAJAR " & username)
            
            tUser = NameIndex(username)
            
            If tUser > 0 Then
                Call ResetFacciones(tUser)
            Else

                If PersonajeExiste(username) Then
                    Call ResetUserFacciones(username)
                Else
                    Call WriteConsoleMsg(UserIndex, "El personaje " & username & " no existe.", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handle the "RemoveCharFromGuild" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleRemoveCharFromGuild(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 12/26/06
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim username   As String

        Dim GuildIndex As Integer
        
        username = Buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call LogGM(.Name, "/RAJARCLAN " & username)
            
            GuildIndex = modGuilds.m_EcharMiembroDeClan(UserIndex, username)
            
            If GuildIndex = 0 Then
                Call WriteConsoleMsg(UserIndex, "No pertenece a ningUn clan o es fundador.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "Expulsado.", FontTypeNames.FONTTYPE_INFO)
                Call SendData(SendTarget.ToGuildMembers, GuildIndex, PrepareMessageConsoleMsg(username & " ha sido expulsado del clan por los administradores del servidor.", FontTypeNames.FONTTYPE_GUILD))

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handle the "RequestCharMail" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleRequestCharMail(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 12/26/06
    'Request user mail
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim username As String

        Dim mail     As String
        
        username = Buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Or .flags.PrivEspecial Then
            If PersonajeExiste(username) Then
                mail = GetUserEmail(username)
                
                Call WriteConsoleMsg(UserIndex, "Last email de " & username & ":" & mail, FontTypeNames.FONTTYPE_INFO)

            End If

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handle the "SystemMessage" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleSystemMessage(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/29/06
    'Send a message to all the users
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim Message As String

        Message = Buffer.ReadASCIIString()
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call LogGM(.Name, "Mensaje de sistema:" & Message)
            
            Call SendData(SendTarget.Toall, 0, PrepareMessageShowMessageBox(Message))

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handle the "SetMOTD" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleSetMOTD(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 03/31/07
    'Set the MOTD
    'Modified by: Juan Martin Sotuyo Dodero (Maraxus)
    '   - Fixed a bug that prevented from properly setting the new number of lines.
    '   - Fixed a bug that caused the player to be kicked.
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim newMOTD           As String

        Dim auxiliaryString() As String

        Dim LoopC             As Long
        
        newMOTD = Buffer.ReadASCIIString()
        auxiliaryString = Split(newMOTD, vbCrLf)
        
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios)) Then
            Call LogGM(.Name, "Ha fijado un nuevo MOTD")
            
            MaxLines = UBound(auxiliaryString()) + 1
            
            ReDim MOTD(1 To MaxLines)
            
            Call WriteVar(ConfigPath & "Motd.ini", "INIT", "NumLines", CStr(MaxLines))
            
            For LoopC = 1 To MaxLines
                Call WriteVar(ConfigPath & "Motd.ini", "Motd", "Line" & CStr(LoopC), auxiliaryString(LoopC - 1))
                
                MOTD(LoopC).texto = auxiliaryString(LoopC - 1)
            Next LoopC
            
            Call WriteConsoleMsg(UserIndex, "Se ha cambiado el MOTD con exito.", FontTypeNames.FONTTYPE_INFO)

        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handle the "ChangeMOTD" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleChangeMOTD(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin sotuyo Dodero (Maraxus)
    'Last Modification: 12/29/06
    'Change the MOTD
    '***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        If (.flags.Privilegios And (PlayerType.RoleMaster Or PlayerType.User Or PlayerType.Consejero Or PlayerType.SemiDios)) Then
            Exit Sub

        End If
        
        Dim auxiliaryString As String

        Dim LoopC           As Long
        
        For LoopC = LBound(MOTD()) To UBound(MOTD())
            auxiliaryString = auxiliaryString & MOTD(LoopC).texto & vbCrLf
        Next LoopC
        
        If Len(auxiliaryString) >= 2 Then
            If Right$(auxiliaryString, 2) = vbCrLf Then
                auxiliaryString = Left$(auxiliaryString, Len(auxiliaryString) - 2)

            End If

        End If
        
        Call WriteShowMOTDEditionForm(UserIndex, auxiliaryString)

    End With

End Sub

''
' Handle the "Ping" message
'
' @param userIndex The index of the user sending the message

Public Sub HandlePing(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Lucas Tavolaro Ortiz (Tavo)
    'Last Modification: 12/24/06
    'Show guilds messages
    '***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        Call WritePong(UserIndex)

    End With

End Sub

''
' Handle the "SetIniVar" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleSetIniVar(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Brian Chaia (BrianPr)
    'Last Modification: 01/23/10 (Marco)
    'Modify server.ini
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 6 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue
        
        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim sLlave As String

        Dim sClave As String

        Dim sValor As String

        'Obtengo los parametros
        sLlave = Buffer.ReadASCIIString()
        sClave = Buffer.ReadASCIIString()
        sValor = Buffer.ReadASCIIString()

        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) Then

            Dim sTmp As String

            'No podemos modificar [INIT]Dioses ni [Dioses]*
            If (UCase$(sLlave) = "INIT" And UCase$(sClave) = "DIOSES") Or UCase$(sLlave) = "DIOSES" Then
                Call WriteConsoleMsg(UserIndex, "No puedes modificar esa informacion desde aqui!", FontTypeNames.FONTTYPE_INFO)
            Else
                'Obtengo el valor segUn llave y clave
                sTmp = GetVar(ConfigPath & "Server.ini", sLlave, sClave)

                'Si obtengo un valor escribo en el server.ini
                If LenB(sTmp) Then
                    Call WriteVar(ConfigPath & "Server.ini", sLlave, sClave, sValor)
                    Call LogGM(.Name, "Modifico en server.ini (" & sLlave & " " & sClave & ") el valor " & sTmp & " por " & sValor)
                    Call WriteConsoleMsg(UserIndex, "Modifico " & sLlave & " " & sClave & " a " & sValor & ". Valor anterior " & sTmp, FontTypeNames.FONTTYPE_INFO)
                Else
                    Call WriteConsoleMsg(UserIndex, "No existe la llave y/o clave", FontTypeNames.FONTTYPE_INFO)

                End If

            End If

        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handle the "CreatePretorianClan" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleCreatePretorianClan(ByVal UserIndex As Integer)
    '***************************************************
    'Author: ZaMa
    'Last Modification: 29/10/2010
    '***************************************************

    On Error GoTo errHandler

    Dim Map   As Integer
    Dim X     As Integer
    Dim Y     As Integer
    Dim Index As Long
    
    With UserList(UserIndex)
        
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Map = .incomingData.ReadByte()
        X = .incomingData.ReadInteger()
        Y = .incomingData.ReadInteger()
        
        ' User Admin?
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) = 0 Then Exit Sub
        
        If Not PRETORIANOS_ACTIVADO Then
            Call WriteConsoleMsg(UserIndex, "Sistema de pretorianos desactivado.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        ' Valid pos?
        If Not InMapBounds(Map, X, Y) Then
            Call WriteConsoleMsg(UserIndex, "Posicion invalida.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        ' Choose pretorian clan index
        If Map = MAPA_PRETORIANO Then
            Index = ePretorianType.Default ' Default clan
        Else
            Index = ePretorianType.Custom ' Custom Clan
        End If
            
        ' Is already active any clan?
        If Not ClanPretoriano(Index).Active Then
            
            If Not ClanPretoriano(Index).SpawnClan(Map, X, Y, Index) Then
                Call WriteConsoleMsg(UserIndex, "La posicion no es apropiada para crear el clan", FontTypeNames.FONTTYPE_INFO)

            End If
        
        Else
            Call WriteConsoleMsg(UserIndex, "El clan pretoriano se encuentra activo en el mapa " & ClanPretoriano(Index).ClanMap & ". Utilice /EliminarPretorianos MAPA y reintente.", FontTypeNames.FONTTYPE_INFO)

        End If
    
    End With

    Exit Sub

errHandler:
    Call LogError("Error en HandleCreatePretorianClan. Error: " & Err.Number & " - " & Err.description)

End Sub

''
' Handle the "CreatePretorianClan" message
'
' @param userIndex The index of the user sending the message

Public Sub HandleDeletePretorianClan(ByVal UserIndex As Integer)
    '***************************************************
    'Author: ZaMa
    'Last Modification: 29/10/2010
    '***************************************************

    On Error GoTo errHandler
    
    Dim Map   As Integer

    Dim Index As Long
    
    With UserList(UserIndex)
        
        'Remove packet ID
        Call .incomingData.ReadByte
        
        Map = .incomingData.ReadInteger()
        
        ' User Admin?
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios) = 0 Then Exit Sub
        
        If Not PRETORIANOS_ACTIVADO Then
            Call WriteConsoleMsg(UserIndex, "Sistema de pretorianos desactivado.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        ' Valid map?
        If Map < 1 Or Map > NumMaps Then
            Call WriteConsoleMsg(UserIndex, "Mapa invalido.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
    
        For Index = 1 To UBound(ClanPretoriano)
         
            ' Search for the clan to be deleted
            If ClanPretoriano(Index).ClanMap = Map Then
                ClanPretoriano(Index).DeleteClan
                Exit For

            End If
        
        Next Index
    
    End With

    Exit Sub

errHandler:
    Call LogError("Error en HandleDeletePretorianClan. Error: " & Err.Number & " - " & Err.description)

End Sub

Private Sub HandleFightSend(ByVal UserIndex As Integer)
    If UserList(UserIndex).incomingData.Length < 5 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler

    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim ListUsers As String
        Dim GldRequired As Long
        Dim Users() As String
        
        ListUsers = Buffer.ReadASCIIString & "-" & .Name
        GldRequired = Buffer.ReadLong
        
        If Len(ListUsers) >= 1 Then
            Users = Split(ListUsers, "-")
                      
            Call Retos.SendFight(UserIndex, GldRequired, Users)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With
    
errHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

Private Sub HandleFightAccept(ByVal UserIndex As Integer)
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
On Error GoTo errHandler

    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte
        
        Dim username As String
        
        username = Buffer.ReadASCIIString
        
        If Len(username) >= 1 Then
            Call Retos.AcceptFight(UserIndex, username)
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With
    
errHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then _
        Err.Raise Error
End Sub

Private Sub HandleCloseGuild(ByVal UserIndex As Integer)
    
    With UserList(UserIndex)
    
        Call .incomingData.ReadByte
        
        Dim i As Long
        Dim PreviousGuildIndex  As Integer
        
        If Not .GuildIndex >= 1 Then
            Call WriteConsoleMsg(UserIndex, "No perteneces a ningun clan.", FONTTYPE_GUILD)
            Exit Sub

        End If

        If guilds(.GuildIndex).Fundador <> .Name Then
            Call WriteConsoleMsg(UserIndex, "No eres lider del clan.", FONTTYPE_GUILD)
            Exit Sub

        End If
        
        'Ya con cambiarle el nombre a "CLAN CERRADO" ya se omite de la lista de clanes enviadas al cliente.
        'Tambien cambiamos "Founder" y "Leader" a "NADIE" sino no te deja fundar otro clan.
        Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & .GuildIndex, "GuildName", "CLAN CERRADO")
        Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & .GuildIndex, "Founder", "NADIE")
        Call WriteVar(App.Path & "\guilds\guildsinfo.inf", "GUILD" & .GuildIndex, "Leader", "NADIE")
        
        PreviousGuildIndex = .GuildIndex
        
        'Obtenemos la lista de miembros del clan.
        Dim GuildMembers() As String
            GuildMembers = guilds(PreviousGuildIndex).GetMemberList()

        For i = 0 To UBound(GuildMembers)
            Call SaveUserGuildIndex(GuildMembers(i), 0)
            Call SaveUserGuildAspirant(GuildMembers(i), 0)
        Next i
        
        'La borramos junto con la lista de solicitudes.
        Call Kill(App.Path & "\Guilds\" & guilds(PreviousGuildIndex).GuildName & "-members.mem")
        Call Kill(App.Path & "\Guilds\" & guilds(PreviousGuildIndex).GuildName & "-solicitudes.sol")
        
        Call SendData(SendTarget.Toall, 0, PrepareMessageConsoleMsg("El Clan " & guilds(.GuildIndex).GuildName & " ha cerrado sus puertas.", FontTypeNames.FONTTYPE_GUILD))
        
    End With

    ' Actualizamos la base de datos de clanes.
    Call modGuilds.LoadGuildsDB
        
    Exit Sub

End Sub

''
' Handles the "Discord" message.
'
' @param    userIndex The index of the user sending the message.

Private Sub HandleDiscord(ByVal UserIndex As Integer)
'***************************************************
'Author: Lucas Daniel Recoaro (Recox)
'Last Modification: 14/07/19 (Recox)
'Manda un mensaje al server para que el mismo lo envie al bot del discord (Recox)
'***************************************************

    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue
        Set Buffer = New clsByteQueue

        Call Buffer.CopyBuffer(.incomingData)
        
        'Remove packet ID
        Call Buffer.ReadByte

        Dim Chat As String
        Chat = Buffer.ReadASCIIString()

        If LenB(Chat) <> 0 Then
            'Analize chat...
            Call Statistics.ParseChat(Chat)
            
            'Aqui solo vamos a hacer un request a los endpoints de la aplicacion en Node.js
            'el repositorio para hacer funcionar esto, es este: https://github.com/ao-libre/ao-api-server
            'Si no tienen interes en usarlo pueden desactivarlo en el Server.ini
            If ConexionAPI Then
                
                                Call ApiEndpointSendCustomCharacterMessageDiscord(Chat, .Name, .Desc)
                Call WriteConsoleMsg(UserIndex, "Link Discord: https://discord.gg/xbAuHcf - El bot de Discord recibio y envio lo siguiente: " & Chat, FontTypeNames.FONTTYPE_INFOBOLD)

            Else
                Call WriteConsoleMsg(UserIndex, "(api - node.js)  El modulo para usar esta funcion no esta instalado en este servidor. http://www.github.com/ao-libre/ao-api-server para mas informacion / more info.", FontTypeNames.FONTTYPE_INFOBOLD)

            End If
        
        End If
        
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

    End With
    
errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

Public Sub HandleLimpiarMundo(ByVal UserIndex As Integer)
'***************************************************
'Author: Jopi
'Last Modification: 11/01/2020
'Fuerza una limpieza del mundo.
'***************************************************
    
    'Remove packet ID
    Call UserList(UserIndex).incomingData.ReadByte
    
    'Me fijo si es GM
    If Not EsGm(UserIndex) Then Exit Sub
    
    
    Call LogGM(UserList(UserIndex).Name, " forzo la limpieza del mundo.")
    
    tickLimpieza = 301
    
End Sub

Public Sub HandleEditGems(ByVal UserIndex As Integer)
'***************************************************
'Author: Lorwik
'Last Modification: 30/04/2020
'Edita las gemas del usuario
'***************************************************
    
    Dim username As String
    Dim CantGems As Long
    Dim Opcion As Byte
    Dim gemasBack As Long
    Dim modificado As Boolean
    
    With UserList(UserIndex)

        'Remove packet ID
        Call .incomingData.ReadByte
        
        username = .incomingData.ReadASCIIString
        CantGems = .incomingData.ReadLong
        Opcion = .incomingData.ReadByte
        
        'Me fijo si es Admin
        If Not EsAdmin(UserList(UserIndex).Name) Then Exit Sub
        
        If username = "" Then
            Call WriteConsoleMsg(UserIndex, "¡Faltan parametros!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If CantGems > 10000 Then
            Call WriteConsoleMsg(UserIndex, "El valor de las Gemas no puede superar 10000", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        gemasBack = .AccountInfo.Gemas
        
        Select Case Opcion
        
            Case 0 'Editar las gemas
                .AccountInfo.Gemas = CantGems
                If Cuentas.SaveAccountGemasDatabase(username, .AccountInfo.Gemas) Then
                    Call WriteConsoleMsg(UserIndex, "Se editaron " & CantGems & " Gemas Winter a la cuenta de " & username, FontTypeNames.FONTTYPE_INFO)
                    modificado = True
                    
                Else
                    Call WriteConsoleMsg(UserIndex, "ERROR: No se pudo editar las gemas a la cuenta del usuario." & username, FontTypeNames.FONTTYPE_INFO)
                    modificado = False
                    
                End If
            
            Case 1 'Sumar las gemas
                .AccountInfo.Gemas = .AccountInfo.Gemas + CantGems
                
                If Cuentas.SaveAccountGemasDatabase(username, .AccountInfo.Gemas) Then
                    Call WriteConsoleMsg(UserIndex, "Se sumaron " & CantGems & " Gemas Winter a la cuenta de " & username & ". Ahora tiene " & Cuentas.GetGemasDatabase(username) & " Gemas Winter. ", FontTypeNames.FONTTYPE_INFO)
                    modificado = True
                    
                Else
                    Call WriteConsoleMsg(UserIndex, "ERROR: No se pudo sumar las gemas a la cuenta del usuario." & username, FontTypeNames.FONTTYPE_INFO)
                    modificado = False
                    
                End If
                
            Case 2 'Restar las gemas
                .AccountInfo.Gemas = .AccountInfo.Gemas - CantGems
            
                If Cuentas.SaveAccountGemasDatabase(username, .AccountInfo.Gemas) Then
                    Call WriteConsoleMsg(UserIndex, "Se restaron " & CantGems & " Gemas Winter a la cuenta de " & username & ". Ahora tiene " & Cuentas.GetGemasDatabase(username) & " Gemas Winter. ", FontTypeNames.FONTTYPE_INFO)
                    modificado = True
                    
                Else
                    Call WriteConsoleMsg(UserIndex, "ERROR: No se pudo restar las gemas de la cuenta del usuario." & username, FontTypeNames.FONTTYPE_INFO)
                    modificado = False
                    
                End If
        End Select
        
        If Not modificado Then _
            .AccountInfo.Gemas = gemasBack
        
    End With
    
End Sub

Public Sub HandleConsultarGemas(ByVal UserIndex As Integer)
'***************************************************
'Author: Lorwik
'Last Modification: 30/04/2020
'Consulta las gemas del usuario
'***************************************************

    Dim username As String
    
    With UserList(UserIndex)
    
        'Remove packet ID
        Call .incomingData.ReadByte
        
        username = .incomingData.ReadASCIIString
    
        'Me fijo si es Admin
        If Not EsAdmin(UserList(UserIndex).Name) Then Exit Sub
        
        If username = "" Then
            Call WriteConsoleMsg(UserIndex, "¡Faltan parametros!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        Call WriteConsoleMsg(UserIndex, username & " tiene " & Cuentas.GetGemasDatabase(username) & " Gemas Winter en su cuenta.", FontTypeNames.FONTTYPE_INFO)
    
    End With
End Sub

''
' Handles the "CraftsmanCreate" message.
'
' @param    UserIndex The index of the user sending the message.

Private Sub HandleCraftsmanCreate(ByVal UserIndex As Integer)
    '***************************************************
    'Author: WyroX
    'Last Modification: 27/01/2020
    '***************************************************

    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    With UserList(UserIndex).incomingData
        'Remove packet ID
        Call .ReadByte

        Dim Item As Integer

        Item = .ReadInteger()
        If Item < LBound(ObjArtesano) Or Item > UBound(ObjArtesano) Then Exit Sub

        If UserList(UserIndex).flags.Muerto = 1 Then
            Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
            Exit Sub
        End If

        'El target es un NPC valido?
        If UserList(UserIndex).flags.TargetNPC = 0 Then Exit Sub

        'Es el artesano?
        If Npclist(UserList(UserIndex).flags.TargetNPC).NPCtype <> eNPCType.Artesano Then Exit Sub

        'Esta cerca?
        If Distancia(Npclist(UserList(UserIndex).flags.TargetNPC).Pos, UserList(UserIndex).Pos) > 3 Then
            Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos del artesano.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        'User retira el item del slot
        Call ArtesanoConstruirItem(UserIndex, Item)

    End With
End Sub

Private Sub HandleChatGlobal(ByVal UserIndex As Integer)
'***************************************************
'Autor: Lorwik
'Fecha: 09/06/2020
'Descripción: Conversaciones por chat global
'***************************************************

    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

On Error GoTo errHandler

    With UserList(UserIndex)
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)
      
        'Remove packet ID
        Call Buffer.ReadByte
      
        Dim Message As String
        Message = Buffer.ReadASCIIString()
      
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
      
        '¿El chat global esta activo?
        If GlobalChatActive = True Then

            '¿Esta muerto?
            If .flags.Muerto = 1 Then
                Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
                Exit Sub
            End If

            '¿Tiene el nivel requerido?
            If .Stats.ELV < MINLVLGLOBAL Then
                Call WriteConsoleMsg(UserIndex, "Para usar el chat global debes ser nivel " & MINLVLGLOBAL & " como minimo.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
            
            '¿El usuario esta silenciado?
            If UserList(UserIndex).flags.Global = 0 Then
                Call WriteConsoleMsg(UserIndex, "No puedes hablar por el chat global por que has sido silenciado.", FontTypeNames.FONTTYPE_INFO)
            Else
                
                'Si no pasaron 5 segundos desde el último mensaje global enviado por el usuario
                If IntervaloPermiteChatGlobal(UserIndex) Then
                    Call SendData(SendTarget.Toall, 0, PrepareMessageConsoleMsg(.Name & "> " & Message, FontTypeNames.FONTTYPE_TALK))
                    Call LogGlobal(.Name & "> " & Message)
    
                Else
                    Call WriteConsoleMsg(UserIndex, "Debes esperar al menos " & INTERVALO_GLOBAL / 1000 & " segundos entre cada mensaje.", FontTypeNames.FONTTYPE_INFO)
                    
                End If
            End If
            
        Else
        
            Call WriteConsoleMsg(UserIndex, "El chat global se encuentra deshabilitado en estos momentos.", FontTypeNames.FONTTYPE_INFO)
        End If
          
    End With

errHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0
  
    'Destroy auxiliar buffer
    Set Buffer = Nothing
  
    If Error <> 0 Then _
        Err.Raise Error
End Sub

Private Sub HandleSilenciarGlobal(ByVal UserIndex As Integer)
'***************************************************
'Autor: Lorwik
'Fecha: 09/06/2020
'Descripción: Silencia a un usuario del chat global
'***************************************************
    If UserList(UserIndex).incomingData.Length < 3 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

On Error GoTo errHandler
    With UserList(UserIndex)
    
        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim username As String
        Dim tUser As Integer

        username = Buffer.ReadASCIIString()

        'Reemplazamos el + con el espacio
        If InStr(1, username, "+") Then
            username = Replace(username, "+", " ")
        End If

        If Not .flags.Privilegios And PlayerType.User Then
            tUser = NameIndex(username)
                
            'Se encuentra offline?
            If tUser <= 0 Then
               Call WriteConsoleMsg(UserIndex, "El personaje no esta online.", FontTypeNames.FONTTYPE_INFO)
               
            Else 'Si esta online...
            
                '¿Tiene el chat global activado?
                If UserList(tUser).flags.Global = 1 Then
                    UserList(tUser).flags.Global = 0
                    Call WriteConsoleMsg(UserIndex, "Se ha silenciado el chat global del usuario: " & UserList(tUser).Name & ".", FontTypeNames.FONTTYPE_INFO)
                    'Le metemos un plus y le avisamos al usuario que se portó mal y que no tiene más chat global
                    Call WriteShowMessageBox(tUser, "Has sido silenciado del chat global indefinidamente.")
                    
                    'Guardamos en el log del gm la acción
                    Call LogGM(.Name, "Ha prohibido el uso del chat global de: " & UserList(tUser).Name)
                    'Guardamos en el log de los usuarios con el chat global prohibido el gm que realizó la acción y el usuario
          
                    Call BanGlobalChatAgregar(UserList(tUser).Name)
          
                    'Flush the other user's buffer
                    Call FlushBuffer(tUser)
                    
                Else '¿Tiene el chat global desactivado?
                
                    'Si el flag era 0 lo restauramos a 1 y le avisamos al gm
                    UserList(tUser).flags.Global = 1
                    Call WriteConsoleMsg(UserIndex, "Has sido des-silenciado del chat global. Utilizalo con moderación: " & UserList(tUser).Name & ".", FontTypeNames.FONTTYPE_INFO)
                    
                    'Guardamos en el log del gm la acción
                    Call LogGM(.Name, "Ha reestablecido el chat global de: " & UserList(tUser).Name)
          
                    Call BanGlobalChatQuitar(UserList(tUser).Name)
                End If
            End If
        End If

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
    End With

errHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then _
        Err.Raise Error
End Sub

Public Sub HandleToggleGlobal(ByVal UserIndex As Integer)
'***************************************************
'Author: MAB
'Declaraciones: Si queres vivir mejor, ponele un IF a tu vida
'***************************************************
On Error GoTo errHandler

With UserList(UserIndex)
    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
    Call Buffer.CopyBuffer(.incomingData)
      
    'Remove packet ID
    Call Buffer.ReadByte
  
    'Solo un Dios o un Admin puede activar/desactivar el global
    If .flags.Privilegios > PlayerType.Dios Or .flags.Privilegios > PlayerType.Admin Then
  
        'Si está activo (que por defecto lo está) entonces lo desactivamos y enviamos un mensaje global a todos los usuarios
        If GlobalChatActive = True Then
            GlobalChatActive = False
            Call SendData(SendTarget.Toall, 0, PrepareMessageConsoleMsg("Servidor> El chat global ha sido desactivado.", FontTypeNames.FONTTYPE_SERVER))
            
        Else
        
            'Si estaba deshabilitado, lo habilitamos e informamos a todos los usuarios
            GlobalChatActive = True
            Call SendData(SendTarget.Toall, 0, PrepareMessageConsoleMsg("Servidor> El chat global fue activado.", FontTypeNames.FONTTYPE_SERVER))
        End If
  
    End If
  
    'If we got here then packet is complete, copy data back to original queue
    Call .incomingData.CopyBuffer(Buffer)
End With

errHandler:
    Dim Error As Long
    Error = Err.Number
On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing
  
    If Error <> 0 Then Err.Raise Error
End Sub

Private Sub HandleAccionInventario(ByVal UserIndex As Integer)
'***********************************
'Autor: Lorwik
'Fecha: 14/07/2020
'Descripcion: Recibimos una accion sobre el inventario ¿Que debemos hacer?
'***********************************
    Dim itemSlot As Byte
    Dim ObjIndex As Long
    
    With UserList(UserIndex)
    
        'Remove packet ID
        .incomingData.ReadByte
        
        itemSlot = .incomingData.ReadByte

        ObjIndex = .Invent.Object(itemSlot).ObjIndex

        'Esta el user muerto y no esta usando una piedra de hogar?
        If .flags.Muerto = 1 And ObjData(ObjIndex).OBJType <> otRunaHogar Then
            Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
            Exit Sub
        End If
        
        If .flags.Comerciando Then Exit Sub
        
        'Validate item slot
        If itemSlot > .CurrentInventorySlots Or itemSlot < 1 Then Exit Sub
        
        If ObjIndex = 0 Then Exit Sub
        
        Select Case ObjData(ObjIndex).OBJType
        
            Case eOBJType.otCasco, eOBJType.otArmadura, eOBJType.otEscudo, eOBJType.otAnillo
                Call EquiparInvItem(UserIndex, itemSlot)
                
            Case eOBJType.otWeapon
                
                If ObjData(ObjIndex).proyectil Then '¿Es un arco?
                    If .Invent.Object(itemSlot).Equipped Then '¿Lo tiene ya equipado?
                        Call UseInvItem(UserIndex, itemSlot) 'Lo usamos
                        Exit Sub
                    End If
                End If
                
                'Equipamos el arma
                Call EquiparInvItem(UserIndex, itemSlot)
            
            Case Else
                Call UseInvItem(UserIndex, itemSlot)
                
        End Select
        
    End With

End Sub

Private Sub HandleInvocar(ByVal UserIndex As Integer)
'***********************************
'Autor: Lorwik
'Fecha: 19/07/2020
'Descripcion: Comienza el rito de invocacion
'***********************************

    'Remove packet ID
    UserList(UserIndex).incomingData.ReadByte
    
    Call IniciarRitoInvocacion(UserIndex)
    
End Sub

Public Sub HandleIniciaSubasta(ByVal UserIndex As Integer)
'***************************************************
'Author: Standelf
'Last Modification: 25/05/2010
'***************************************************
    With UserList(UserIndex).incomingData
        'Remove Packet ID
        Call .ReadByte
        Call Iniciar_Subasta(UserIndex, .ReadInteger, .ReadInteger, .ReadLong)
    End With
End Sub

Public Sub HandleCancelarSubasta(ByVal UserIndex As Integer)
'***************************************************
'Author: Lorwik
'Last Modification: 19/08/2020
'Descripción: El user no subasta y cierra el form
'***************************************************
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        .flags.Subastando = False
    End With
End Sub
 
Public Sub HandleConsultarSubasta(ByVal UserIndex As Integer)
'***************************************************
'Author: Standelf
'Last Modification: 05/07/2010
'***************************************************
    With UserList(UserIndex).incomingData
        'Remove Packet ID
        Call .ReadByte
        Call Consultar_Subasta(UserIndex)
    End With
End Sub
 
Public Sub HandleOfertaSubasta(ByVal UserIndex As Integer)
'***************************************************
'Author: Standelf
'Last Modification: 25/05/2010
'***************************************************
    With UserList(UserIndex).incomingData
        'Remove Packet ID
        Call .ReadByte
        Call Ofertar_Subasta(UserIndex, .ReadLong)
    End With
End Sub

Public Sub HandleRespuestaInstruccion(ByVal UserIndex As Integer)
'***************************************************
'Autor: Lorwik
'Fecha: 19/08/2020
'Descripcion: Recibe respuesta de la instruccion
'***************************************************
    Dim Respuesta As Boolean
    
    With UserList(UserIndex)
        'Remove Packet ID
        Call .incomingData.ReadByte
        
        Respuesta = .incomingData.ReadBoolean
        
        If Respuesta Then
            Call AccionProfesion(UserIndex)
            
        Else
            'Reseteamos los flags
            .flags.ProfInstruyendo = 0
            .flags.Instruyendo = 0
            
        End If
        
    End With
    
End Sub

Public Sub HandleMsgAmigo(ByVal UserIndex As Integer)

    If UserList(UserIndex).incomingData.Length < 3 Then
        Call Err.Raise(UserList(UserIndex).incomingData.NotEnoughDataErrCode)
        Exit Sub
    End If

    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim Mensaje As String
        Dim i       As Long

        Mensaje = Buffer.ReadASCIIString()

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

        For i = 1 To MAXAMIGOS

            If .Amigos(i).Index > 0 Then
                Call WriteConsoleMsg(.Amigos(i).Index, "FMSG[" & .Name & "]: " & Mensaje, FontTypeNames.FONTTYPE_GM)
            End If

        Next i

        Call WriteConsoleMsg(UserIndex, "FMSG[" & .Name & "]: " & Mensaje, FontTypeNames.FONTTYPE_GM)

    End With

errHandler:

    Dim Error As Long
        Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then Call Err.Raise(Error)
End Sub

Public Sub HandleOnAmigo(ByVal UserIndex As Integer)
'***********************************
'Autor: ???
'Fecha: ???
'Descripcion: ¿Amigos conectados?
'***********************************

    With UserList(UserIndex)

        'Remove packet ID
        Call .incomingData.ReadByte
        Dim list As String
        Dim i    As Long

        For i = 1 To MAXAMIGOS

            If .Amigos(i).Index > 0 Then
                list = list & "[" & UserList(.Amigos(i).Index).Name & "-" & MapZonas(UserList(.Amigos(i).Index).Pos.Map, UserZonaId(UserIndex)).Name & "];"
            End If

        Next i

        If LenB(list) > 0 Then
            Call WriteConsoleMsg(UserIndex, "Onlines: " & list, FontTypeNames.FONTTYPE_CONSEJO)
        Else
            Call WriteConsoleMsg(UserIndex, "No tienes ningun amigo conectado.", FontTypeNames.FONTTYPE_GM)
        End If

    End With

End Sub

Public Sub HandleAddAmigo(ByVal UserIndex As Integer)
'***********************************
'Autor: ???
'Fecha: ???
'Descripcion: Recibe una peticion para agregar un amigo
'***********************************

    If UserList(UserIndex).incomingData.Length < 3 Then
        Call Err.Raise(UserList(UserIndex).incomingData.NotEnoughDataErrCode)
        Exit Sub
    End If

    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As clsByteQueue: Set Buffer = New clsByteQueue
        Call Buffer.CopyBuffer(.incomingData)

        'Remove packet ID
        Call Buffer.ReadByte

        Dim username  As String
        Dim tUserName As String
        Dim caso      As Byte
        Dim razon     As String
        Dim tUser     As Integer
        Dim Slot      As Byte

        username = Buffer.ReadASCIIString()
        caso = Buffer.ReadByte
        tUser = NameIndex(username)

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)

        'Mandar solicitudad de amistad
        If caso = 1 Then

            If AgregarAmigo(UserIndex, tUser, razon) = True Then
                Call WriteConsoleMsg(UserIndex, "Se ha enviado una solicitud de amistad a " & UserList(tUser).Name, FontTypeNames.FONTTYPE_CONSEJO)
                Call WriteConsoleMsg(tUser, UserList(UserIndex).Name & " quiere ser tu amigo. Para aceptarlo usa el comando /FADD " & .Name, FontTypeNames.FONTTYPE_CONSEJO)
                UserList(tUser).Quien = .Name

            Else
                Call WriteConsoleMsg(UserIndex, razon, FontTypeNames.FONTTYPE_CONSEJO)

            End If
            'Confirmar solicitudad de amistad

        ElseIf caso > 1 Then

            If AgregarAmigo(UserIndex, tUser, razon) = True Then

                If LenB(.Quien) >= 3 Then

                    If UCase$(.Quien) = UCase$(UserList(tUser).Name) Then

                        Slot = BuscarSlotAmigoVacio(UserIndex)

                        .Amigos(Slot).Nombre = UserList(tUser).Name
                        .Amigos(Slot).Ignorado = 0

                        Call ActualizarSlotAmigo(UserIndex, Slot)

                        Slot = BuscarSlotAmigoVacio(tUser)

                        UserList(tUser).Amigos(Slot).Nombre = .Name
                        UserList(tUser).Amigos(Slot).Ignorado = 0

                        Call ActualizarSlotAmigo(tUser, Slot)

                        Call WriteConsoleMsg(UserIndex, UserList(tUser).Name & " agregado", FontTypeNames.FONTTYPE_DIOS)

                        Call WriteConsoleMsg(tUser, .Name & " agregado", FontTypeNames.FONTTYPE_DIOS)

                        Slot = ObtenerIndexLibre(UserIndex)

                        If Slot > 0 Then
                            .Amigos(Slot).Index = tUser
                        End If

                        Slot = ObtenerIndexLibre(tUser)

                        If Slot > 0 Then
                            UserList(tUser).Amigos(Slot).Index = UserIndex
                        End If

                        .Quien = vbNullString

                    Else
                        Call WriteConsoleMsg(UserIndex, "Solicitud de amistad invalida.", FontTypeNames.FONTTYPE_CONSEJO)

                    End If

                End If

            Else
                Call WriteConsoleMsg(UserIndex, razon, FontTypeNames.FONTTYPE_CONSEJO)

            End If

        End If

    End With

errHandler:

    Dim Error As Long
        Error = Err.Number

    On Error GoTo 0

    'Destroy auxiliar buffer
    Set Buffer = Nothing

    If Error <> 0 Then Call Err.Raise(Error)

End Sub

Public Sub HandleDelAmigo(ByVal UserIndex As Integer)
'***********************************
'Autor: ???
'Fecha: ???
'Descripcion: Recibe una peticion para eliminar un amigo
'***********************************

    With UserList(UserIndex)

        'Remove packet ID
        Call .incomingData.ReadByte

        Dim Slot     As Byte
        Dim tUser    As Integer
        Dim username As String

        Slot = .incomingData.ReadByte()

        If Slot <= 0 Or Slot > MAXAMIGOS Then Exit Sub

        'Por las duditas :P
        If LenB(.Amigos(Slot).Nombre) = 0 Then Exit Sub

        tUser = NameIndex(.Amigos(Slot).Nombre)
        username = .Amigos(Slot).Nombre

        Call WriteConsoleMsg(UserIndex, .Amigos(Slot).Nombre & " ha sido borrado de la lista de amigos.", FontTypeNames.FONTTYPE_GMMSG)

        'reseteamos el slot
        .Amigos(Slot).Nombre = vbNullString
        .Amigos(Slot).Ignorado = 0
        Call ActualizarSlotAmigo(UserIndex, Slot)

        If tUser > 0 Then

            'Puede pasar....
            If BuscarSlotAmigoName(tUser, .Name) Then

                Call WriteConsoleMsg(tUser, .Name & "te ha borrado de su lista de amigos.", FontTypeNames.FONTTYPE_GMMSG)

                Slot = BuscarSlotAmigoNameSlot(tUser, .Name)

                UserList(tUser).Amigos(Slot).Ignorado = 0
                UserList(tUser).Amigos(Slot).Nombre = vbNullString

                Call ActualizarSlotAmigo(tUser, Slot)

                Slot = ObtenerIndexUsuado(UserIndex, tUser)

                If Slot > 0 Then
                    .Amigos(Slot).Index = 0
                End If

                Slot = ObtenerIndexUsuado(tUser, UserIndex)

                If Slot > 0 Then
                    UserList(tUser).Amigos(Slot).Index = 0
                End If

            End If

        Else

            'verificamos desde el char
            Call BorrarAmigo(username, .Name)

        End If

    End With

End Sub

Public Sub HandleQuestListRequest(ByVal UserIndex As Integer)
    '****************************************************
    'Autor: Amraphen
    'Fecha: 30/01/2010
    'Descripcion: Maneja el paquete QuestListRequest.
    '****************************************************
 
    'Leemos el paquete
    Call UserList(UserIndex).incomingData.ReadByte
    
    Call WriteQuestListSend(UserIndex)

End Sub

Public Sub HandleQuestDetailsRequest(ByVal UserIndex As Integer)
    '****************************************************
    'Autor: Amraphen
    'Fecha: 30/01/2010
    'Descripcion: Maneja el paquete QuestInfoRequest.
    '****************************************************
    
    Dim Questslot As Byte
 
    With UserList(UserIndex)
        'Leemos el paquete
        Call .incomingData.ReadByte
        
        Questslot = .incomingData.ReadByte

        Call WriteQuestDetails(UserIndex, .QuestStats.Quests(.QuestStats.QuestEnCurso(Questslot)).QuestIndex, Questslot)
    End With

End Sub
 
Public Sub HandleQuestAbandon(ByVal UserIndex As Integer)
    '****************************************************
    'Autor: Lorwik
    'Fecha: 04/05/2021
    'Descripcion: El usuario quiere abandonar una quest
    '****************************************************
    
    Dim Questslot As Integer
    
    With UserList(UserIndex)
    
        'Leemos el paquete.
        Call .incomingData.ReadByte
        
        Questslot = .incomingData.ReadByte
        
        Call modQuests.userAbandonaQuest(UserIndex, Questslot)
        
    End With
    
End Sub

Public Sub HandleQuestAccept(ByVal UserIndex As Integer)
    '****************************************************
    'Autor: Lorwik
    'Fecha: 04/05/2021
    'Descripcion: El usuario quiere aceptar una quest
    '****************************************************
 
    Call UserList(UserIndex).incomingData.ReadByte
 
    Call modQuests.userAceptaquest(UserIndex)

End Sub
 
 Public Sub HandleQuest(ByVal UserIndex As Integer)

    '****************************************************
    'Maneja el paquete Quest.
    'Last modified: 18/05/2020
    'Lorwik: Paso todo el chequeo y la accion a otro sub refractorio
    '****************************************************
    
    Dim NPCIndex As Integer

    'Leemos el paquete
    Call UserList(UserIndex).incomingData.ReadByte
 
    NPCIndex = UserList(UserIndex).flags.TargetNPC
    
    
    Call accionUseraNPCQuest(UserIndex, NPCIndex)

End Sub

Private Sub HandleBanSerial(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Lorwik
    'Last Modification: 05/05/2021
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Set Buffer = New clsByteQueue
        
        Call Buffer.CopyBuffer(.incomingData)
  
        'Remove packet ID
        Call Buffer.ReadByte
  
        Dim username As String
   
        username = Buffer.ReadASCIIString()
  
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
  
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
        
            If EsAdmin(username) = False Or EsDios(username) = False Then _
                Call BanSerialOK(UserIndex, username)

        End If

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

Private Sub HandleUnBanSerial(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Lorwik
    'Last Modification: 05/05/2021
    '
    '***************************************************
    If UserList(UserIndex).incomingData.Length < 4 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Set Buffer = New clsByteQueue
        
        Call Buffer.CopyBuffer(.incomingData)
  
        'Remove packet ID
        Call Buffer.ReadByte
  
        Dim username As String
   
        username = Buffer.ReadASCIIString()
                
        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
                
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            Call UnBanSerialOK(UserIndex, username)
      
        End If

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

Private Sub HandleBanTemporal(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Lorwik
    'Last Modification: 05/05/2021
    '
    '***************************************************

    If UserList(UserIndex).incomingData.Length < 6 Then
        Err.Raise UserList(UserIndex).incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errHandler

    With UserList(UserIndex)

        'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
        Dim Buffer As New clsByteQueue
        Set Buffer = New clsByteQueue
        
        Call Buffer.CopyBuffer(.incomingData)
  
        'Remove packet ID
        Call Buffer.ReadByte
  
        Dim username As String

        Dim Reason   As String

        Dim Dias     As Byte
  
        username = Buffer.ReadASCIIString()
        Reason = Buffer.ReadASCIIString()
        Dias = Buffer.ReadByte()

        'If we got here then packet is complete, copy data back to original queue
        Call .incomingData.CopyBuffer(Buffer)
  
        If (Not .flags.Privilegios And PlayerType.RoleMaster) <> 0 And (.flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios)) <> 0 Then
            If EsAdmin(username) = False Or EsDios(username) = False Then
                Call Admin.BanCharacter(UserIndex, username, Reason, Dias)
                
            End If

        End If

    End With

errHandler:

    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set Buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

Private Sub HandleShopInit(ByVal UserIndex As Integer)
    '****************************************************
    'Autor: Lorwik
    'Fecha: 16/05/2022
    'Descripcion: El usuario quiere abrir la Shop
    '****************************************************
 
    Call UserList(UserIndex).incomingData.ReadByte
 
    Call WriteMostrarShop(UserIndex)
End Sub

Private Sub HandleBuyShop(ByVal UserIndex As Integer)
    '****************************************************
    'Autor: Lorwik
    'Fecha: 16/05/2022
    'Descripcion: El usuario quiere abrir la Shop
    '****************************************************
 
    Dim Objeto As Integer
    Dim gemasBack As Long
 
    With UserList(UserIndex)
    
        Call .incomingData.ReadByte
        
        Objeto = .incomingData.ReadInteger
                
        If Objeto < 1 Then
            Call WriteConsoleMsg(UserIndex, "Selecciona un item.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If ShopObject(Objeto).Valor > .AccountInfo.Gemas Then
            Call WriteConsoleMsg(UserIndex, "No tienes gemas suficiente para comprar ese producto.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Me fijo si tiene espacio en el inventario
        Dim objInventario As obj
        
        objInventario.Amount = ShopObject(Objeto).Amount
        objInventario.ObjIndex = ShopObject(Objeto).ObjIndex
        
        If Not MeterItemEnInventario(UserIndex, objInventario) Then
            Call WriteConsoleMsg(UserIndex, "No tienes espacio en tu inventario.", FontTypeNames.FONTTYPE_INFO)
            
        Else
        
            gemasBack = .AccountInfo.Gemas
            'Restamos las gemas
            .AccountInfo.Gemas = .AccountInfo.Gemas - ShopObject(Objeto).Valor
            
            If Cuentas.SaveAccountGemasDatabase(.Name, .AccountInfo.Gemas) Then
                'Guardamos el log de la transacción
                Call LogShopTransactions(.Name & " | Compró -> " & ObjData(objInventario.ObjIndex).Name & " | Valor -> " & ShopObject(Objeto).Valor & " | Gemas Actuales -> " & .AccountInfo.Gemas)
                Call WriteActualizarGemasShop(UserIndex)
                    
            Else
                Call WriteConsoleMsg(UserIndex, "ERROR: No se pudo restar las gemas de la cuenta del usuario." & .Name, FontTypeNames.FONTTYPE_INFO)
                'Transacción no realizada, devolvemos las gemas
                .AccountInfo.Gemas = gemasBack
                    
            End If
        
        End If
    
    End With
 
End Sub

Private Sub HandleInitPVP(ByVal UserIndex As Integer)
    '****************************************************
    'Autor: Lorwik
    'Fecha: 22/05/2022
    'Descripcion: El usuario quiere abrir el PVP
    '****************************************************
    
    Call UserList(UserIndex).incomingData.ReadByte
    
    Call WriteMostrarPVP(UserIndex)
    
End Sub

Public Sub HandleDueloSet(ByVal UserIndex As Integer)
    '****************************************************
    'Autor: Lorwik
    'Fecha: 27/05/2022
    'Descripcion: El usuario quiere Duelos
    '****************************************************
    Dim TipoDuelo As Byte
    Dim MapaDuelo As Byte
    With UserList(UserIndex)
        Call .incomingData.ReadByte
        
        TipoDuelo = .incomingData.ReadByte
        Select Case TipoDuelo
            Case 0
                Call EsperarOponenteDuelo(UserIndex)
            Case 1 'Duelo Clasico sin ELO
                Call EsperarOponenteDueloClasico(UserIndex, False)
            Case 2 'Duelo Clasico con ELO
                Call EsperarOponenteDueloClasico(UserIndex, True)
            Case 3 'Arena de Rinkel
                Call modArenaRinkel.EntrarArenaRinkel(UserIndex)
            Case 50 'Comenzar Arena de Rinkel
                Call modArenaRinkel.Preparar(UserIndex)
            Case Else
                Exit Sub
        End Select
    End With
End Sub
