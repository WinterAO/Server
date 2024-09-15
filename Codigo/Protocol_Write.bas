Attribute VB_Name = "Protocol_Write"
Option Explicit

''
'When we have a list of strings, we use this to separate them and prevent
'having too many string lengths in the queue. Yes, each string is NULL-terminated :P
Private Const SEPARATOR As String * 1 = vbNullChar

''
'Auxiliar ByteQueue used as buffer to generate messages not intended to be sent right away.
'Specially usefull to create a message once and send it over to several clients.
Private auxiliarBuffer  As clsByteQueue

Private Enum ServerPacketID
    Logged = 1                  ' LOGGED
    RemoveDialogs               ' QTDL
    RemoveCharDialog            ' QDL
    NavigateToggle              ' NAVEG
    Disconnect                  ' FINOK
    CommerceEnd                 ' FINCOMOK
    BankEnd                     ' FINBANOK
    CommerceInit                ' INITCOM
    BankInit                    ' INITBANCO
    CommerceChat
    UpdateSta                   ' ASS
    UpdateMana                  ' ASM
    UpdateHP                    ' ASH
    UpdateGold                  ' ASG
    UpdateBankGold
    UpdateExp                   ' ASE
    ChangeMap                   ' CM
    PosUpdate                   ' PU
    ChatOverHead                ' ||
    ConsoleMsg                  ' || - Beware!! its the same as above, but it was properly splitted
    ScreenMsg
    GuildChat                   ' |+
    ShowMessageBox              ' !!
    UserIndexInServer           ' IU
    UserCharIndexInServer       ' IP
    CharacterCreate             ' CC
    CharacterRemove             ' BP
    CharacterChangeNick
    CharacterMove               ' MP, +, * and _ '
    ForceCharMove
    CharacterChange             ' CP
    HeadingChange
    ObjectCreate                ' HO
    ObjectDelete                ' BO
    BlockPosition               ' BQ
    UserCommerceInit            ' INITCOMUSU
    UserCommerceEnd             ' FINCOMUSUOK
    UserOfferConfirm
    PlayMusic                    ' TM
    PlayWave                     ' TW
    guildList                    ' GL
    AreaChanged                  ' CA
    PauseToggle                  ' BKW
    ActualizarClima
    CreateFX                     ' CFX
    UpdateUserStats              ' EST
    ChangeInventorySlot          ' CSI
    ChangeBankSlot               ' SBO
    ChangeSpellSlot              ' SHS
    Atributes                    ' ATR
    InitTrabajo
    RestOK                       ' DOK
    errorMsg                     ' ERR
    Blind                        ' CEGU
    Dumb                         ' DUMB
    ShowSignal                   ' MCAR
    ChangeNPCInventorySlot       ' NPCI
    UpdateHungerAndThirst        ' EHYS
    Fame                         ' FAMA
    MiniStats                    ' MEST
    AddForumMsg                  ' FMSG
    ShowForumForm                ' MFOR
    SetInvisible                 ' NOVER
    MeditateToggle               ' MEDOK
    BlindNoMore                  ' NSEGUE
    DumbNoMore                   ' NESTUP
    SendSkills                   ' SKILLS
    TrainerCreatureList          ' LSTCRI
    guildNews                    ' GUILDNE
    OfferDetails                 ' PEACEDE & ALLIEDE
    AlianceProposalsList         ' ALLIEPR
    PeaceProposalsList           ' PEACEPR
    CharacterInfo                ' CHRINFO
    GuildLeaderInfo              ' LEADERI
    GuildMemberInfo
    GuildDetails                 ' CLANDET
    ShowGuildFundationForm       ' SHOWFUN
    ParalizeOK                   ' PARADOK
    ShowUserRequest              ' PETICIO
    ChangeUserTradeSlot          ' COMUSUINV
    SendNight                    ' NOC
    Pong
    UpdateTagAndStatus
    BattleGs                     'Battlegrounds
    MostrarShop
    ActualizarGemasShop
    SpeedToChar
    
    'GM =  messages
    SpawnList                    ' SPL
    ShowSOSForm                  ' MSOS
    ShowMOTDEditionForm          ' ZMOTD
    ShowGMPanelForm              ' ABPANEL
    UserNameList                 ' LISTUSU
    ShowDenounces
    RecordList
    RecordDetails
    ShowGuildAlign
    ShowPartyForm
    PeticionInvitarParty
    UpdateStrenghtAndDexterity
    UpdateStrenght
    UpdateDexterity
    AddSlots
    MultiMessage
    StopWorking
    CancelOfferItem
    PlayAttackAnim
    FXtoMap
    EnviarPJUserAccount
    SearchList
    QuestDetails
    QuestListSend
    ActualizarNPCQuest
    CreateDamage                ' CDMG
    UserInEvent
    DeletedChar
    EquitandoToggle
    InitCraftman
    EnviarListDeAmigos
    proyectil
    CharParticle
    IniciarSubastaConsulta
    ConfirmarInstruccion
    SetSpeed
    AtaqueNPC
    MostrarPVP
    eBarFx
    ePrivilegios
End Enum

Public Sub InitAuxiliarBuffer()
    '***************************************************
    'Author: ZaMa
    'Last Modification: 15/03/2011
    'Initializaes Auxiliar Buffer
    '***************************************************
    Set auxiliarBuffer = New clsByteQueue

End Sub

''
' Writes the "Logged" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLoggedMessage(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Logged" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    With UserList(UserIndex)
        Call .outgoingData.WriteByte(ServerPacketID.Logged)
    
        Call .outgoingData.WriteByte(.clase)

    End With
    
    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "RemoveDialogs" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRemoveAllDialogs(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RemoveDialogs" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.RemoveDialogs)
    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "RemoveCharDialog" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character whose dialog will be removed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRemoveCharDialog(ByVal UserIndex As Integer, ByVal CharIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RemoveCharDialog" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageRemoveCharDialog(CharIndex))
    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "MessageCreateDamage" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character whose dialog will be removed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMessageCreateDamage(ByVal UserIndex As Integer, ByVal dano As Long, ByVal Damage_Type As Byte)

    '***************************************************
    'Author: Lorwik
    'Fecha: 22/06/2020
    'Writes the "MessageCreateDamage" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    With UserList(UserIndex)
    
        Call .outgoingData.WriteASCIIStringFixed(PrepareMessageCreateDamage(.Pos.X, .Pos.Y, dano, Damage_Type))
    
    End With
    
    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "NavigateToggle" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNavigateToggle(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "NavigateToggle" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.NavigateToggle)
    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "Disconnect" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDisconnect(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Disconnect" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.Disconnect)
    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "UserOfferConfirm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserOfferConfirm(ByVal UserIndex As Integer)

    '***************************************************
    'Author: ZaMa
    'Last Modification: 14/12/2009
    'Writes the "UserOfferConfirm" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.UserOfferConfirm)
    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "CommerceEnd" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceEnd(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CommerceEnd" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.CommerceEnd)
    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "BankEnd" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankEnd(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BankEnd" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.BankEnd)
    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "CommerceInit" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceInit(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CommerceInit" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.CommerceInit)
    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "BankInit" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankInit(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BankInit" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.BankInit)
    Call UserList(UserIndex).outgoingData.WriteLong(UserList(UserIndex).Stats.Banco)
    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "UserCommerceInit" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceInit(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UserCommerceInit" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.UserCommerceInit)
    Call UserList(UserIndex).outgoingData.WriteASCIIString(UserList(UserIndex).ComUsu.DestNick)
    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "UserCommerceEnd" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceEnd(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UserCommerceEnd" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.UserCommerceEnd)
    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "UpdateSta" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateSta(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UpdateMana" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateSta)
        Call .WriteInteger(UserList(UserIndex).Stats.MinSta)

    End With

    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "UpdateMana" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateMana(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UpdateMana" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateMana)
        Call .WriteInteger(UserList(UserIndex).Stats.MinMAN)

    End With

    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "UpdateHP" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateHP(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UpdateMana" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateHP)
        Call .WriteInteger(UserList(UserIndex).Stats.MinHp)

    End With

    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "UpdateGold" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateGold(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UpdateGold" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateGold)
        Call .WriteLong(UserList(UserIndex).Stats.Gld)

    End With

    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "UpdateBankGold" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateBankGold(ByVal UserIndex As Integer)

    '***************************************************
    'Author: ZaMa
    'Last Modification: 14/12/2009
    'Writes the "UpdateBankGold" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateBankGold)
        Call .WriteLong(UserList(UserIndex).Stats.Banco)

    End With

    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "UpdateExp" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateExp(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UpdateExp" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateExp)
        Call .WriteLong(UserList(UserIndex).Stats.Exp)

    End With

    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "UpdateStrenghtAndDexterity" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateStrenghtAndDexterity(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Budi
    'Last Modification: 11/26/09
    'Writes the "UpdateStrenghtAndDexterity" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateStrenghtAndDexterity)
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza))
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))

    End With

    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

' Writes the "UpdateStrenghtAndDexterity" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateDexterity(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Budi
    'Last Modification: 11/26/09
    'Writes the "UpdateStrenghtAndDexterity" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateDexterity)
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))

    End With

    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

' Writes the "UpdateStrenghtAndDexterity" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateStrenght(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Budi
    'Last Modification: 11/26/09
    'Writes the "UpdateStrenghtAndDexterity" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateStrenght)
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza))

    End With

    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ChangeMap" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    map The new map to load.
' @param    version The version of the map in the server to check if client is properly updated.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMap(ByVal UserIndex As Integer, _
                          ByVal Map As Integer, _
                          ByVal version As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ChangeMap" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeMap)
        Call .WriteInteger(Map)
        Call .WriteInteger(version)

    End With

    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "PosUpdate" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePosUpdate(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "PosUpdate" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.PosUpdate)
        Call .WriteInteger(UserList(UserIndex).Pos.X)
        Call .WriteInteger(UserList(UserIndex).Pos.Y)

    End With

    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ChatOverHead" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Chat Text to be displayed over the char's head.
' @param    CharIndex The character uppon which the chat will be displayed.
' @param    Color The color to be used when displaying the chat.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChatOverHead(ByVal UserIndex As Integer, _
                             ByVal Chat As String, _
                             ByVal CharIndex As Integer, _
                             ByVal r As Byte, _
                             ByVal g As Byte, _
                             ByVal b As Byte, _
                             Optional ByVal NoConsole As Boolean = False)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ChatOverHead" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageChatOverHead(Chat, CharIndex, r, g, b, NoConsole))
    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ConsoleMsg" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Chat Text to be displayed over the char's head.
' @param    FontIndex Index of the FONTTYPE structure to use.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteConsoleMsg(ByVal UserIndex As Integer, _
                           ByVal Chat As String, _
                           ByVal FontIndex As FontTypeNames)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ConsoleMsg" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageConsoleMsg(Chat, FontIndex))
    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ScreenMsg" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Chat Text to be displayed over the char's head.
' @param    FontIndex Index of the FONTTYPE structure to use.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteScreenMsg(ByVal UserIndex As Integer, _
                           ByVal msgPrimario As String, _
                           ByVal msgScundario As String)

    '***************************************************
    'Author: Lorwik
    'Last Modification: 24/04/2021
    'Writes the "ScreenMsg" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageScreenMsg(msgPrimario, msgScundario))
    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Public Sub WriteCommerceChat(ByVal UserIndex As Integer, _
                             ByVal Chat As String, _
                             ByVal FontIndex As FontTypeNames)

    '***************************************************
    'Author: ZaMa
    'Last Modification: 05/17/06
    'Writes the "ConsoleMsg" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareCommerceConsoleMsg(Chat, FontIndex))
    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub
            
''
' Writes the "GuildChat" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Chat Text to be displayed over the char's head.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildChat(ByVal UserIndex As Integer, ByVal Chat As String)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildChat" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageGuildChat(Chat))
    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ShowMessageBox" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Message Text to be displayed in the message box.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteShowMessageBox(ByVal UserIndex As Integer, ByVal Message As String)
'***************************************************
'Author: Juan Martin Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ShowMessageBox" message to the given user's outgoing data buffer
'***************************************************
    
        On Error GoTo errHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ShowMessageBox)
        Call .WriteASCIIString(Message)

    End With

    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "UserIndexInServer" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserIndexInServer(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UserIndexInServer" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UserIndexInServer)
        Call .WriteInteger(UserIndex)

    End With

    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "UserCharIndexInServer" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCharIndexInServer(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UserIndexInServer" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UserCharIndexInServer)
        Call .WriteInteger(UserList(UserIndex).Char.CharIndex)

    End With

    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "CharacterCreate" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    body Body index of the new character.
' @param    head Head index of the new character.
' @param    heading Heading in which the new character is looking.
' @param    CharIndex The index of the new character.
' @param    X X coord of the new character's position.
' @param    Y Y coord of the new character's position.
' @param    weapon Weapon index of the new character.
' @param    shield Shield index of the new character.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @param    helmet Helmet index of the new character.
' @param    name Name of the new character.
' @param    criminal Determines if the character is a criminal or not.
' @param    privileges Sets if the character is a normal one or any kind of administrative character.
' @param    NoShadow establece si el cuerpo no emite sombra
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharacterCreate(ByVal UserIndex As Integer, _
                                ByVal body As Integer, _
                                ByVal Head As Integer, _
                                ByVal Heading As eHeading, _
                                ByVal CharIndex As Integer, _
                                ByVal X As Integer, _
                                ByVal Y As Integer, _
                                ByVal weapon As Integer, _
                                ByVal shield As Integer, _
                                ByVal FX As Integer, _
                                ByVal FXLoops As Integer, _
                                ByVal helmet As Integer, _
                                ByVal AnimAtaque As Integer, _
                                ByVal Name As String, _
                                ByVal NickColor As Byte, _
                                ByVal Privileges As Byte, _
                                ByVal GrhAura As Long, _
                                ByVal AuraColor As Long, _
                                ByVal speeding As Single, _
                                Optional ByVal estadoQuest As Byte = 255)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CharacterCreate" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageCharacterCreate(body, Head, Heading, CharIndex, X, Y, weapon, shield, FX, FXLoops, helmet, AnimAtaque, Name, NickColor, Privileges, GrhAura, AuraColor, speeding, estadoQuest))
    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "CharacterRemove" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character to be removed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharacterRemove(ByVal UserIndex As Integer, ByVal CharIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CharacterRemove" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageCharacterRemove(CharIndex))
    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "CharacterMove" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character which is moving.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharacterMove(ByVal UserIndex As Integer, _
                              ByVal CharIndex As Integer, _
                              ByVal X As Integer, _
                              ByVal Y As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CharacterMove" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageCharacterMove(CharIndex, X, Y))
    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Public Sub WriteForceCharMove(ByVal UserIndex, ByVal Direccion As eHeading)

    '***************************************************
    'Author: ZaMa
    'Last Modification: 26/03/2009
    'Writes the "ForceCharMove" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageForceCharMove(Direccion))
    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "CharacterChange" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    body Body index of the new character.
' @param    head Head index of the new character.
' @param    heading Heading in which the new character is looking.
' @param    CharIndex The index of the new character.
' @param    weapon Weapon index of the new character.
' @param    shield Shield index of the new character.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @param    helmet Helmet index of the new character.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharacterChange(ByVal UserIndex As Integer, _
                                ByVal body As Integer, _
                                ByVal Head As Integer, _
                                ByVal Heading As eHeading, _
                                ByVal CharIndex As Integer, _
                                ByVal weapon As Integer, _
                                ByVal shield As Integer, _
                                ByVal FX As Integer, _
                                ByVal FXLoops As Integer, _
                                ByVal helmet As Integer, _
                                ByVal AuraAnim As Long, _
                                ByVal AuraColor As Long, _
                                Optional ByVal QuestStatus As Byte = 255)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CharacterChange" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageCharacterChange(body, Head, Heading, CharIndex, weapon, shield, FX, FXLoops, helmet, AuraAnim, AuraColor, QuestStatus))
    Exit Sub
    
errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub



''
' Writes the "ObjectCreate" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    GrhIndex Grh of the object.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @param    Shadow establece si el objeto emite sombra
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteObjectCreate(ByVal UserIndex As Integer, _
                             ByVal GrhIndex As Long, _
                             ByVal ParticulaIndex As Integer, _
                             ByVal X As Integer, _
                             ByVal Y As Integer, _
                             Optional ByVal Shadow As Byte = 0)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ObjectCreate" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageObjectCreate(GrhIndex, ParticulaIndex, X, Y, Shadow))
    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ObjectDelete" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteObjectDelete(ByVal UserIndex As Integer, ByVal X As Integer, ByVal Y As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ObjectDelete" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageObjectDelete(X, Y))
    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "BlockPosition" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @param    Blocked True if the position is blocked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlockPosition(ByVal UserIndex As Integer, _
                              ByVal X As Integer, _
                              ByVal Y As Integer, _
                              ByVal Blocked As Boolean)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BlockPosition" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.BlockPosition)
        Call .WriteInteger(X)
        Call .WriteInteger(Y)
        Call .WriteBoolean(Blocked)

    End With

    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "PlayMidi" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    midi The midi to be played.
' @param    loops Number of repets for the midi.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePlayMusic(ByVal UserIndex As Integer, _
                         ByVal music As Integer, _
                         Optional ByVal loops As Integer = -1)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "PlayMidi" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessagePlayMusic(music, loops))
    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "PlayWave" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    wave The wave to be played.
' @param    X The X position in map coordinates from where the sound comes.
' @param    Y The Y position in map coordinates from where the sound comes.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePlayWave(ByVal UserIndex As Integer, _
                         ByVal wave As Byte, _
                         ByVal X As Integer, _
                         ByVal Y As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 08/08/07
    'Last Modified by: Rapsodius
    'Added X and Y positions for 3D Sounds
    '***************************************************
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessagePlayWave(wave, X, Y))
    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "GuildList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    GuildList List of guilds to be sent.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildList(ByVal UserIndex As Integer, ByRef guildList() As String)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildList" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Dim Tmp As String

    Dim i   As Long
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.guildList)
        
        ' Prepare guild name's list
        For i = LBound(guildList()) To UBound(guildList())
            Tmp = Tmp & guildList(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)

    End With

    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "AreaChanged" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAreaChanged(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "AreaChanged" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.AreaChanged)
        Call .WriteInteger(UserList(UserIndex).Pos.X)
        Call .WriteInteger(UserList(UserIndex).Pos.Y)
        Call .WriteByte(UserList(UserIndex).Char.Heading)

    End With

    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "PauseToggle" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePauseToggle(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "PauseToggle" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessagePauseToggle())
    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ActualizarClima" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteActualizarClima(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ActualizarClima" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageActualizarClima())
    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "CreateFX" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character upon which the FX will be created.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreateFX(ByVal UserIndex As Integer, _
                         ByVal CharIndex As Integer, _
                         ByVal FX As Integer, _
                         ByVal FXLoops As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CreateFX" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageCreateFX(CharIndex, FX, FXLoops))
    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "UpdateUserStats" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateUserStats(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UpdateUserStats" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateUserStats)
        Call .WriteInteger(UserList(UserIndex).Stats.MaxHp)
        Call .WriteInteger(UserList(UserIndex).Stats.MinHp)
        Call .WriteInteger(UserList(UserIndex).Stats.MaxMAN)
        Call .WriteInteger(UserList(UserIndex).Stats.MinMAN)
        Call .WriteInteger(UserList(UserIndex).Stats.MaxSta)
        Call .WriteInteger(UserList(UserIndex).Stats.MinSta)
        Call .WriteLong(UserList(UserIndex).Stats.Gld)
        Call .WriteByte(UserList(UserIndex).Stats.ELV)
        Call .WriteLong(UserList(UserIndex).Stats.ELU)
        Call .WriteLong(UserList(UserIndex).Stats.Exp)

    End With

    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ChangeInventorySlot" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    slot Inventory slot which needs to be updated.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeInventorySlot(ByVal UserIndex As Integer, ByVal Slot As Byte)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 3/12/09
    'Writes the "ChangeInventorySlot" message to the given user's outgoing data buffer
    '3/12/09: Budi - Ahora se envia MaxDef y MinDef en lugar de Def
    '***************************************************
    On Error GoTo errHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeInventorySlot)
        Call .WriteByte(Slot)
        
        Dim ObjIndex As Integer

        Dim obData   As ObjData
        
        ObjIndex = UserList(UserIndex).Invent.Object(Slot).ObjIndex
        
        If ObjIndex > 0 Then
            obData = ObjData(ObjIndex)

        End If
        
        Call .WriteInteger(ObjIndex)
        Call .WriteASCIIString(obData.Name)
        Call .WriteInteger(UserList(UserIndex).Invent.Object(Slot).Amount)
        Call .WriteBoolean(UserList(UserIndex).Invent.Object(Slot).Equipped)
        Call .WriteLong(obData.GrhIndex)
        Call .WriteByte(obData.OBJType)
        Call .WriteInteger(obData.MaxHIT)
        Call .WriteInteger(obData.MinHIT)
        Call .WriteInteger(obData.MaxDef)
        Call .WriteInteger(obData.MinDef)
        Call .WriteSingle(SalePrice(ObjIndex))
        Call .WriteBoolean(ItemNoUsaConUser(UserIndex, ObjIndex))

    End With

    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Public Sub WriteAddSlots(ByVal UserIndex As Integer, ByVal Mochila As eMochilas)

    '***************************************************
    'Author: Budi
    'Last Modification: 01/12/09
    'Writes the "AddSlots" message to the given user's outgoing data buffer
    '***************************************************
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.AddSlots)
        Call .WriteByte(Mochila)

    End With

End Sub

''
' Writes the "ChangeBankSlot" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    slot Inventory slot which needs to be updated.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeBankSlot(ByVal UserIndex As Integer, ByVal Slot As Byte)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 12/03/09
    'Writes the "ChangeBankSlot" message to the given user's outgoing data buffer
    '12/03/09: Budi - Ahora se envia MaxDef y MinDef en lugar de solo Def
    '***************************************************
    On Error GoTo errHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeBankSlot)
        Call .WriteByte(Slot)
        
        Dim ObjIndex As Integer

        Dim obData   As ObjData
        
        ObjIndex = UserList(UserIndex).BancoInvent.Object(Slot).ObjIndex
        
        Call .WriteInteger(ObjIndex)
        
        If ObjIndex > 0 Then
            obData = ObjData(ObjIndex)

        End If
        
        Call .WriteASCIIString(obData.Name)
        Call .WriteInteger(UserList(UserIndex).BancoInvent.Object(Slot).Amount)
        Call .WriteLong(obData.GrhIndex)
        Call .WriteByte(obData.OBJType)
        Call .WriteInteger(obData.MaxHIT)
        Call .WriteInteger(obData.MinHIT)
        Call .WriteInteger(obData.MaxDef)
        Call .WriteInteger(obData.MinDef)
        Call .WriteLong(obData.Valor)
        Call .WriteBoolean(ItemNoUsaConUser(UserIndex, ObjIndex))

    End With

    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ChangeSpellSlot" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    slot Spell slot to update.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeSpellSlot(ByVal UserIndex As Integer, ByVal Slot As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 05/17/06
'Writes the "ChangeSpellSlot" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeSpellSlot)
        Call .WriteByte(Slot)
        Call .WriteInteger(UserList(UserIndex).Stats.UserHechizos(Slot))
        
        If UserList(UserIndex).Stats.UserHechizos(Slot) > 0 Then
            Call .WriteASCIIString(Hechizos(UserList(UserIndex).Stats.UserHechizos(Slot)).Nombre)
        Else
            Call .WriteASCIIString("(None)")
        End If
    End With
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

''
' Writes the "Atributes" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAttributes(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Atributes" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.Atributes)
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Fuerza))
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Agilidad))
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Inteligencia))
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma))
        Call .WriteByte(UserList(UserIndex).Stats.UserAtributos(eAtributos.Constitucion))

    End With

    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "BlacksmithWeapons" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteInitTrabajo(ByVal UserIndex As Integer, ByVal Profesion As Byte)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 04/15/2008 (NicoNZ) Habia un error al fijarse los skills del personaje
    'Writes the "BlacksmithWeapons" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Dim i                           As Long
    Dim j                           As Byte
    Dim obj(1 To MAXUSERRECETAS)    As Long
    Dim validIndexes()              As Integer
    Dim Count                       As Integer
    Dim SlotProfesion               As Integer
    Dim PrecioConstruccion          As Long
    
    With UserList(UserIndex)
    
        'Obtenemos el slot de la profesion
        SlotProfesion = ConoceProfesion(UserIndex, Profesion)
        
        'Si por un casual nos llegara que no la conoce...
        If SlotProfesion < 0 Then Exit Sub
    
        Call .outgoingData.WriteByte(ServerPacketID.InitTrabajo)
        Call .outgoingData.WriteByte(Profesion) 'Mando la profesion
        
        For i = 1 To MAXUSERRECETAS

            ' Can the user create this object? If so add it to the list....
            If .Profesion(SlotProfesion).Recetas(i) > 0 Then
                Count = Count + 1
                obj(Count) = .Profesion(SlotProfesion).Recetas(i)

            End If

        Next i
        
        ' Write the number of objects in the list
        Call .outgoingData.WriteInteger(Count)
        
        ' Write the needed data of each object
        For i = 1 To Count
            Call .outgoingData.WriteASCIIString(ObjData(obj(i)).Name)
            Call .outgoingData.WriteLong(ObjData(obj(i)).GrhIndex)
            
            Select Case Profesion
            
                Case eSkill.Alquimia
                    PrecioConstruccion = ObjData(obj(i)).SkAlquimia * 3000
                    
                Case eSkill.herreria
                    PrecioConstruccion = ObjData(obj(i)).SkHerreria * 3000
                    
                Case eSkill.Carpinteria
                    PrecioConstruccion = ObjData(obj(i)).SkCarpinteria * 3000
            
            End Select
            
            If .AccountInfo.esVIP = True Then _
                PrecioConstruccion = PrecioConstruccion / 2
            
            Call .outgoingData.WriteLong(PrecioConstruccion)
            
            For j = 1 To MAXMATERIALES
                If ObjData(obj(i)).Materiales(j) > 0 Then
                    Call .outgoingData.WriteLong(ObjData(ObjData(obj(i)).Materiales(j)).GrhIndex)
                    Call .outgoingData.WriteInteger(ObjData(obj(i)).CantMateriales(j))
                    Call .outgoingData.WriteASCIIString(ObjData(ObjData(obj(i)).Materiales(j)).Name)
                    
                Else
                    Call .outgoingData.WriteLong(0)
                    Call .outgoingData.WriteInteger(0)
                    Call .outgoingData.WriteASCIIString("Nada")
                    
                End If
            Next j
            
            Call .outgoingData.WriteInteger(obj(i))
        Next i
        
        .flags.Trabajando = Profesion

    End With

    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Public Sub WriteInitCraftsman(ByVal UserIndex As Integer)
    '***************************************************
    'Author: WyroX
    'Last Modification: 27/01/2020
    'Writes the "InitCraftman" message to the given user's outgoing data buffer
    '***************************************************

    On Error GoTo errHandler

    Dim i              As Long
    Dim j              As Long
    Dim obj            As ObjData
    Dim ObjRequired    As ObjData

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.InitCraftman)
        
        ' Write cost of crafting
        Call .WriteLong(ArtesaniaCosto)
        
        ' Write the number of objects in the list
        Call .WriteInteger(UBound(ObjArtesano))
        
        ' Write the needed data of each object
        For i = 1 To UBound(ObjArtesano)
            obj = ObjData(ObjArtesano(i))
            Call .WriteASCIIString(obj.Name)
            Call .WriteLong(obj.GrhIndex)
            Call .WriteInteger(ObjArtesano(i))

            Call .WriteByte(UBound(obj.ItemCrafteo))
            
            For j = 1 To UBound(obj.ItemCrafteo)
                ObjRequired = ObjData(obj.ItemCrafteo(j).ObjIndex)
                Call .WriteASCIIString(ObjRequired.Name)
                Call .WriteLong(ObjRequired.GrhIndex)
                Call .WriteInteger(obj.ItemCrafteo(j).ObjIndex)
                Call .WriteInteger(obj.ItemCrafteo(j).Amount)
            Next j
        Next i
        
    End With

    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If
End Sub

''
' Writes the "RestOK" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRestOK(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RestOK" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.RestOK)
    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ErrorMsg" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    message The error message to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteErrorMsg(ByVal UserIndex As Integer, ByVal Message As String)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ErrorMsg" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageErrorMsg(Message))

    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "Blind" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlind(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Blind" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.Blind)
    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "Dumb" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDumb(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Dumb" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.Dumb)
    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ShowSignal" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    objIndex Index of the signal to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowSignal(ByVal UserIndex As Integer, ByVal ObjIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowSignal" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ShowSignal)
        Call .WriteASCIIString(ObjData(ObjIndex).texto)
        Call .WriteLong(ObjData(ObjIndex).GrhSecundario)

    End With

    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ChangeNPCInventorySlot" message to the given user's outgoing data buffer.
'
' @param    UserIndex   User to which the message is intended.
' @param    slot        The inventory slot in which this item is to be placed.
' @param    obj         The object to be set in the NPC's inventory window.
' @param    price       The value the NPC asks for the object.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeNPCInventorySlot(ByVal UserIndex As Integer, _
                                       ByVal Slot As Byte, _
                                       ByRef obj As obj, _
                                       ByVal price As Single)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 12/03/09
    'Last Modified by: Budi
    'Writes the "ChangeNPCInventorySlot" message to the given user's outgoing data buffer
    '12/03/09: Budi - Ahora se envia MaxDef y MinDef en lugar de solo Def
    '***************************************************
    On Error GoTo errHandler

    Dim ObjInfo As ObjData
    
    If obj.ObjIndex >= LBound(ObjData()) And obj.ObjIndex <= UBound(ObjData()) Then
        ObjInfo = ObjData(obj.ObjIndex)

    End If
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeNPCInventorySlot)
        Call .WriteByte(Slot)
        Call .WriteASCIIString(ObjInfo.Name)
        Call .WriteInteger(obj.Amount)
        Call .WriteSingle(price)
        Call .WriteLong(ObjInfo.GrhIndex)
        Call .WriteInteger(obj.ObjIndex)
        Call .WriteByte(ObjInfo.OBJType)
        Call .WriteInteger(ObjInfo.MaxHIT)
        Call .WriteInteger(ObjInfo.MinHIT)
        Call .WriteInteger(ObjInfo.MaxDef)
        Call .WriteInteger(ObjInfo.MinDef)
        Call .WriteBoolean(ItemNoUsaConUser(UserIndex, obj.ObjIndex))

    End With

    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "UpdateHungerAndThirst" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpdateHungerAndThirst(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UpdateHungerAndThirst" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UpdateHungerAndThirst)
        Call .WriteByte(UserList(UserIndex).Stats.MaxAGU)
        Call .WriteByte(UserList(UserIndex).Stats.MinAGU)
        Call .WriteByte(UserList(UserIndex).Stats.MaxHam)
        Call .WriteByte(UserList(UserIndex).Stats.MinHam)

    End With

    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "Fame" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteFame(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Fame" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.Fame)
        
        Call .WriteLong(UserList(UserIndex).Reputacion.AsesinoRep)
        Call .WriteLong(UserList(UserIndex).Reputacion.BandidoRep)
        Call .WriteLong(UserList(UserIndex).Reputacion.BurguesRep)
        Call .WriteLong(UserList(UserIndex).Reputacion.LadronesRep)
        Call .WriteLong(UserList(UserIndex).Reputacion.NobleRep)
        Call .WriteLong(UserList(UserIndex).Reputacion.PlebeRep)
        Call .WriteLong(UserList(UserIndex).Reputacion.Promedio)

    End With

    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "MiniStats" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMiniStats(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "MiniStats" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.MiniStats)
        
        Call .WriteLong(UserList(UserIndex).Faccion.CiudadanosMatados)
        Call .WriteLong(UserList(UserIndex).Faccion.CriminalesMatados)
        
        'TODO : Este valor es calculable, no deberia NI EXISTIR, ya sea en el servidor ni en el cliente!!!
        Call .WriteLong(UserList(UserIndex).Stats.UsuariosMatados)
        
        Call .WriteInteger(UserList(UserIndex).Stats.NPCsMuertos)
        
        Call .WriteByte(UserList(UserIndex).clase)
        Call .WriteByte(UserList(UserIndex).Raza)
        
        Call .WriteLong(UserList(UserIndex).Counters.Pena)

    End With

    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "AddForumMsg" message to the given user's outgoing data buffer.
'
' @param    title The title of the message to display.
' @param    message The message to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAddForumMsg(ByVal UserIndex As Integer, _
                            ByVal ForumType As eForumType, _
                            ByRef Title As String, _
                            ByRef Author As String, _
                            ByRef Message As String)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 02/01/2010
    'Writes the "AddForumMsg" message to the given user's outgoing data buffer
    '02/01/2010: ZaMa - Now sends Author and forum type
    '***************************************************
    On Error GoTo errHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.AddForumMsg)
        Call .WriteByte(ForumType)
        Call .WriteASCIIString(Title)
        Call .WriteASCIIString(Author)
        Call .WriteASCIIString(Message)

    End With

    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ShowForumForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowForumForm(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowForumForm" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Dim Visibilidad   As Byte

    Dim CanMakeSticky As Byte
    
    With UserList(UserIndex)
        Call .outgoingData.WriteByte(ServerPacketID.ShowForumForm)
        
        Visibilidad = eForumVisibility.ieGENERAL_MEMBER
        
        If esCaos(UserIndex) Or EsGm(UserIndex) Then
            Visibilidad = Visibilidad Or eForumVisibility.ieCAOS_MEMBER

        End If
        
        If esArmada(UserIndex) Or EsGm(UserIndex) Then
            Visibilidad = Visibilidad Or eForumVisibility.ieREAL_MEMBER

        End If
        
        Call .outgoingData.WriteByte(Visibilidad)
        
        ' Pueden mandar sticky los gms o los del consejo de armada/caos
        If EsGm(UserIndex) Then
            CanMakeSticky = 2
        ElseIf (.flags.Privilegios And PlayerType.ChaosCouncil) <> 0 Then
            CanMakeSticky = 1
        ElseIf (.flags.Privilegios And PlayerType.RoyalCouncil) <> 0 Then
            CanMakeSticky = 1

        End If
        
        Call .outgoingData.WriteByte(CanMakeSticky)

    End With

    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "SetInvisible" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex The char turning visible / invisible.
' @param    invisible True if the char is no longer visible, False otherwise.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSetInvisible(ByVal UserIndex As Integer, _
                             ByVal CharIndex As Integer, _
                             ByVal invisible As Boolean)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SetInvisible" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(PrepareMessageSetInvisible(CharIndex, invisible))
    
    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "MeditateToggle" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMeditateToggle(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "MeditateToggle" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.MeditateToggle)
    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "BlindNoMore" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBlindNoMore(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BlindNoMore" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.BlindNoMore)
    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "DumbNoMore" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDumbNoMore(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "DumbNoMore" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.DumbNoMore)
    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "SendSkills" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSendSkills(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 11/19/09
    'Writes the "SendSkills" message to the given user's outgoing data buffer
    '11/19/09: Pato - Now send the percentage of progress of the skills.
    '***************************************************
    On Error GoTo errHandler

    Dim i As Long
    
    With UserList(UserIndex)
        Call .outgoingData.WriteByte(ServerPacketID.SendSkills)
        Call .outgoingData.WriteByte(.clase)
        
        For i = 1 To NUMSKILLS
            Call .outgoingData.WriteByte(UserList(UserIndex).Stats.UserSkills(i))
        Next i

    End With

    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "TrainerCreatureList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    npcIndex The index of the requested trainer.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTrainerCreatureList(ByVal UserIndex As Integer, ByVal NPCIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "TrainerCreatureList" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Dim i   As Long

    Dim str As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.TrainerCreatureList)
        
        For i = 1 To Npclist(NPCIndex).NroCriaturas
            str = str & Npclist(NPCIndex).Criaturas(i).NpcName & SEPARATOR
        Next i
        
        If LenB(str) > 0 Then str = Left$(str, Len(str) - 1)
        
        Call .WriteASCIIString(str)

    End With

    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "GuildNews" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    guildNews The guild's news.
' @param    enemies The list of the guild's enemies.
' @param    allies The list of the guild's allies.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildNews(ByVal UserIndex As Integer, _
                          ByVal guildNews As String, _
                          ByRef enemies() As String, _
                          ByRef allies() As String)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildNews" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Dim i   As Long

    Dim Tmp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.guildNews)
        
        Call .WriteASCIIString(guildNews)
        
        'Prepare enemies' list
        For i = LBound(enemies()) To UBound(enemies())
            Tmp = Tmp & enemies(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
        
        Tmp = vbNullString

        'Prepare allies' list
        For i = LBound(allies()) To UBound(allies())
            Tmp = Tmp & allies(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)

    End With

    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "OfferDetails" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    details Th details of the Peace proposition.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOfferDetails(ByVal UserIndex As Integer, ByVal details As String)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "OfferDetails" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Dim i As Long
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.OfferDetails)
        
        Call .WriteASCIIString(details)

    End With

    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "AlianceProposalsList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    guilds The list of guilds which propossed an alliance.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAlianceProposalsList(ByVal UserIndex As Integer, ByRef guilds() As String)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "AlianceProposalsList" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Dim i   As Long

    Dim Tmp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.AlianceProposalsList)
        
        ' Prepare guild's list
        For i = LBound(guilds()) To UBound(guilds())
            Tmp = Tmp & guilds(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)

    End With

    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "PeaceProposalsList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    guilds The list of guilds which propossed peace.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePeaceProposalsList(ByVal UserIndex As Integer, ByRef guilds() As String)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "PeaceProposalsList" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Dim i   As Long

    Dim Tmp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.PeaceProposalsList)
                
        ' Prepare guilds' list
        For i = LBound(guilds()) To UBound(guilds())
            Tmp = Tmp & guilds(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)

    End With

    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "CharacterInfo" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    charName The requested char's name.
' @param    race The requested char's race.
' @param    class The requested char's class.
' @param    gender The requested char's gender.
' @param    level The requested char's level.
' @param    gold The requested char's gold.
' @param    reputation The requested char's reputation.
' @param    previousPetitions The requested char's previous petitions to enter guilds.
' @param    currentGuild The requested char's current guild.
' @param    previousGuilds The requested char's previous guilds.
' @param    RoyalArmy True if tha char belongs to the Royal Army.
' @param    CaosLegion True if tha char belongs to the Caos Legion.
' @param    citicensKilled The number of citicens killed by the requested char.
' @param    criminalsKilled The number of criminals killed by the requested char.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCharacterInfo(ByVal UserIndex As Integer, ByVal CharName As String, ByVal race As eRaza, ByVal Class As eClass, ByVal gender As eGenero, ByVal level As Byte, ByVal Gold As Long, ByVal bank As Long, ByVal reputation As Long, ByVal previousPetitions As String, ByVal currentGuild As String, ByVal previousGuilds As String, ByVal RoyalArmy As Boolean, ByVal CaosLegion As Boolean, ByVal citicensKilled As Long, ByVal criminalsKilled As Long)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CharacterInfo" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.CharacterInfo)
        
        Call .WriteASCIIString(CharName)
        Call .WriteByte(race)
        Call .WriteByte(Class)
        Call .WriteByte(gender)
        
        Call .WriteByte(level)
        Call .WriteLong(Gold)
        Call .WriteLong(bank)
        Call .WriteLong(reputation)
        
        Call .WriteASCIIString(previousPetitions)
        Call .WriteASCIIString(currentGuild)
        Call .WriteASCIIString(previousGuilds)
        
        Call .WriteBoolean(RoyalArmy)
        Call .WriteBoolean(CaosLegion)
        
        Call .WriteLong(citicensKilled)
        Call .WriteLong(criminalsKilled)

    End With

    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "GuildLeaderInfo" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    guildList The list of guild names.
' @param    memberList The list of the guild's members.
' @param    guildNews The guild's news.
' @param    joinRequests The list of chars which requested to join the clan.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildLeaderInfo(ByVal UserIndex As Integer, _
                                ByRef guildList() As String, _
                                ByRef MemberList() As String, _
                                ByVal guildNews As String, _
                                ByRef joinRequests() As String)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildLeaderInfo" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Dim i   As Long

    Dim Tmp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.GuildLeaderInfo)
        
        ' Prepare guild name's list
        For i = LBound(guildList()) To UBound(guildList())
            Tmp = Tmp & guildList(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
        
        ' Prepare guild member's list
        Tmp = vbNullString

        For i = LBound(MemberList()) To UBound(MemberList())
            Tmp = Tmp & MemberList(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
        
        ' Store guild news
        Call .WriteASCIIString(guildNews)
        
        ' Prepare the join request's list
        Tmp = vbNullString

        For i = LBound(joinRequests()) To UBound(joinRequests())
            Tmp = Tmp & joinRequests(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)

    End With

    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "GuildLeaderInfo" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    guildList The list of guild names.
' @param    memberList The list of the guild's members.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildMemberInfo(ByVal UserIndex As Integer, _
                                ByRef guildList() As String, _
                                ByRef MemberList() As String)

    '***************************************************
    'Author: ZaMa
    'Last Modification: 21/02/2010
    'Writes the "GuildMemberInfo" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Dim i   As Long

    Dim Tmp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.GuildMemberInfo)
        
        ' Prepare guild name's list
        For i = LBound(guildList()) To UBound(guildList())
            Tmp = Tmp & guildList(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)
        
        ' Prepare guild member's list
        Tmp = vbNullString

        For i = LBound(MemberList()) To UBound(MemberList())
            Tmp = Tmp & MemberList(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)

    End With

    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "GuildDetails" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    guildName The requested guild's name.
' @param    founder The requested guild's founder.
' @param    foundationDate The requested guild's foundation date.
' @param    leader The requested guild's current leader.
' @param    URL The requested guild's website.
' @param    memberCount The requested guild's member count.
' @param    electionsOpen True if the clan is electing it's new leader.
' @param    alignment The requested guild's alignment.
' @param    enemiesCount The requested guild's enemy count.
' @param    alliesCount The requested guild's ally count.
' @param    antifactionPoints The requested guild's number of antifaction acts commited.
' @param    codex The requested guild's codex.
' @param    guildDesc The requested guild's description.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildDetails(ByVal UserIndex As Integer, _
                             ByVal GuildName As String, _
                             ByVal founder As String, _
                             ByVal foundationDate As String, _
                             ByVal leader As String, _
                             ByVal URL As String, _
                             ByVal memberCount As Integer, _
                             ByVal electionsOpen As Boolean, _
                             ByVal alignment As String, _
                             ByVal enemiesCount As Integer, _
                             ByVal AlliesCount As Integer, _
                             ByVal antifactionPoints As String, _
                             ByRef codex() As String, _
                             ByVal guildDesc As String)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildDetails" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Dim i    As Long

    Dim temp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.GuildDetails)
        
        Call .WriteASCIIString(GuildName)
        Call .WriteASCIIString(founder)
        Call .WriteASCIIString(foundationDate)
        Call .WriteASCIIString(leader)
        Call .WriteASCIIString(URL)
        
        Call .WriteInteger(memberCount)
        Call .WriteBoolean(electionsOpen)
        
        Call .WriteASCIIString(alignment)
        
        Call .WriteInteger(enemiesCount)
        Call .WriteInteger(AlliesCount)
        
        Call .WriteASCIIString(antifactionPoints)
        
        For i = LBound(codex()) To UBound(codex())
            temp = temp & codex(i) & SEPARATOR
        Next i
        
        If Len(temp) > 1 Then temp = Left$(temp, Len(temp) - 1)
        
        Call .WriteASCIIString(temp)
        
        Call .WriteASCIIString(guildDesc)

    End With

    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ShowGuildAlign" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowGuildAlign(ByVal UserIndex As Integer)

    '***************************************************
    'Author: ZaMa
    'Last Modification: 14/12/2009
    'Writes the "ShowGuildAlign" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.ShowGuildAlign)
    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ShowGuildFundationForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowGuildFundationForm(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowGuildFundationForm" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.ShowGuildFundationForm)
    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ParalizeOK" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteParalizeOK(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 08/12/07
    'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
    'Writes the "ParalizeOK" message to the given user's outgoing data buffer
    'And updates user position
    '***************************************************
    On Error GoTo errHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ParalizeOK)
    End With
    
    Call WritePosUpdate(UserIndex)

    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ShowUserRequest" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    details DEtails of the char's request.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowUserRequest(ByVal UserIndex As Integer, ByVal details As String)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowUserRequest" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ShowUserRequest)
        
        Call .WriteASCIIString(details)

    End With

    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ChangeUserTradeSlot" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    ObjIndex The object's index.
' @param    amount The number of objects offered.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeUserTradeSlot(ByVal UserIndex As Integer, _
                                    ByVal OfferSlot As Byte, _
                                    ByVal ObjIndex As Integer, _
                                    ByVal Amount As Long)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 12/03/09
    'Writes the "ChangeUserTradeSlot" message to the given user's outgoing data buffer
    '25/11/2009: ZaMa - Now sends the specific offer slot to be modified.
    '12/03/09: Budi - Ahora se envia MaxDef y MinDef en lugar de solo Def
    '***************************************************
    On Error GoTo errHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ChangeUserTradeSlot)
        
        Call .WriteByte(OfferSlot)
        Call .WriteInteger(ObjIndex)
        Call .WriteLong(Amount)
        
        If ObjIndex > 0 Then
        
            Call .WriteLong(ObjData(ObjIndex).GrhIndex)
            Call .WriteByte(ObjData(ObjIndex).OBJType)
            Call .WriteInteger(ObjData(ObjIndex).MaxHIT)
            Call .WriteInteger(ObjData(ObjIndex).MinHIT)
            Call .WriteInteger(ObjData(ObjIndex).MaxDef)
            Call .WriteInteger(ObjData(ObjIndex).MinDef)
            Call .WriteLong(SalePrice(ObjIndex))
            Call .WriteASCIIString(ObjData(ObjIndex).Name)
            Call .WriteBoolean(ItemNoUsaConUser(UserIndex, ObjIndex))
            
        Else ' Borra el item
        
            Call .WriteLong(0)
            Call .WriteByte(0)
            Call .WriteInteger(0)
            Call .WriteInteger(0)
            Call .WriteInteger(0)
            Call .WriteInteger(0)
            Call .WriteLong(0)
            Call .WriteASCIIString("")
            Call .WriteBoolean(False)

        End If

    End With

    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "SendNight" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSendNight(ByVal UserIndex As Integer, ByVal night As Boolean)

    '***************************************************
    'Author: Fredy Horacio Treboux (liquid)
    'Last Modification: 01/08/07
    'Writes the "SendNight" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.SendNight)
        Call .WriteBoolean(night)

    End With

    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "SpawnList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    npcNames The names of the creatures that can be spawned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSpawnList(ByVal UserIndex As Integer, ByRef npcNames() As String)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SpawnList" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Dim i   As Long

    Dim Tmp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.SpawnList)
        
        For i = LBound(npcNames()) To UBound(npcNames())
            Tmp = Tmp & npcNames(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)

    End With

    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ShowSOSForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowSOSForm(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowSOSForm" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Dim i   As Long

    Dim Tmp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ShowSOSForm)
        
        For i = 1 To Ayuda.Longitud
            Tmp = Tmp & Ayuda.VerElemento(i) & SEPARATOR
        Next i
        
        If LenB(Tmp) <> 0 Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)

    End With

    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ShowDenounces" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowDenounces(ByVal UserIndex As Integer)

    '***************************************************
    'Author: ZaMa
    'Last Modification: 14/11/2010
    'Writes the "ShowDenounces" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler
    
    Dim DenounceIndex As Long

    Dim DenounceList  As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ShowDenounces)
        
        For DenounceIndex = 1 To Denuncias.Longitud
            DenounceList = DenounceList & Denuncias.VerElemento(DenounceIndex, False) & SEPARATOR
        Next DenounceIndex
        
        If LenB(DenounceList) <> 0 Then DenounceList = Left$(DenounceList, Len(DenounceList) - 1)
        
        Call .WriteASCIIString(DenounceList)

    End With

    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ShowSOSForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowPartyForm(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Budi
    'Last Modification: 11/26/09
    'Writes the "ShowPartyForm" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Dim i                         As Long

    Dim Tmp                       As String

    Dim PI                        As Integer

    Dim members(PARTY_MAXMEMBERS) As Integer
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ShowPartyForm)
        
        PI = UserList(UserIndex).PartyIndex
        Call .WriteByte(CByte(Parties(PI).EsPartyLeader(UserIndex)))
        
        If PI > 0 Then
            Call Parties(PI).ObtenerMiembrosOnline(members())

            For i = 1 To PARTY_MAXMEMBERS

                If members(i) > 0 Then
                    Tmp = Tmp & UserList(members(i)).Name & " (" & Fix(Parties(PI).MiExperiencia(members(i))) & ")" & SEPARATOR

                End If

            Next i

        End If
        
        If LenB(Tmp) <> 0 Then Tmp = Left$(Tmp, Len(Tmp) - 1)
            
        Call .WriteASCIIString(Tmp)
        Call .WriteLong(Parties(PI).ObtenerExperienciaTotal)

    End With

    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Public Sub WritePeticionInvitarParty(ByVal UserIndex As Integer)
   '***************************************************
    'Author: Lorwik
    'Last Modification: 05/11/2020
    '***************************************************
    On Error GoTo errHandler
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.PeticionInvitarParty)

    End With

    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If
End Sub

''
' Writes the "ShowMOTDEditionForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    currentMOTD The current Message Of The Day.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowMOTDEditionForm(ByVal UserIndex As Integer, _
                                    ByVal currentMOTD As String)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowMOTDEditionForm" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.ShowMOTDEditionForm)
        
        Call .WriteASCIIString(currentMOTD)

    End With

    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "ShowGMPanelForm" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowGMPanelForm(ByVal UserIndex As Integer, ByVal ID As Byte)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowGMPanelForm" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    With UserList(UserIndex).outgoingData
    
        Call .WriteByte(ServerPacketID.ShowGMPanelForm)
        Call .WriteByte(ID)
    
    End With

    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "UserNameList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    userNameList List of user names.
' @param    Cant Number of names to send.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserNameList(ByVal UserIndex As Integer, _
                             ByRef userNamesList() As String, _
                             ByVal cant As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06 NIGO:
    'Writes the "UserNameList" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Dim i   As Long

    Dim Tmp As String
    
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.UserNameList)
        
        ' Prepare user's names list
        For i = 1 To cant
            Tmp = Tmp & userNamesList(i) & SEPARATOR
        Next i
        
        If Len(Tmp) Then Tmp = Left$(Tmp, Len(Tmp) - 1)
        
        Call .WriteASCIIString(Tmp)

    End With

    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "Pong" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePong(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Pong" message to the given user's outgoing data buffer
    '***************************************************
    On Error GoTo errHandler

    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.Pong)
    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Flushes the outgoing data buffer of the user.
'
' @param    UserIndex User whose outgoing data buffer will be flushed.

Public Sub FlushBuffer(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Sends all data existing in the buffer
    '***************************************************
    Dim sndData As String
    
    With UserList(UserIndex).outgoingData

        If .Length = 0 Then Exit Sub
        
        sndData = .ReadASCIIStringFixed(.Length)

    End With
    
    ' Tratamos de enviar los datos.
    Dim ret As Long: ret = WsApiEnviar(UserIndex, sndData)
    
    ' Si recibimos un error como respuesta de la API, cerramos el socket.
    If ret <> 0 And ret <> WSAEWOULDBLOCK Then
        ' Close the socket avoiding any critical error
        Call CloseSocketSL(UserIndex)
        Call Cerrar_Usuario(UserIndex)
    End If

End Sub

''
' Prepares the "SetInvisible" message and returns it.
'
' @param    CharIndex The char turning visible / invisible.
' @param    invisible True if the char is no longer visible, False otherwise.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The message is written to no outgoing buffer, but only prepared in a single string to be easily sent to several clients.

Public Function PrepareMessageSetInvisible(ByVal CharIndex As Integer, _
                                           ByVal invisible As Boolean) As String

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "SetInvisible" message and returns it.
    '***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.SetInvisible)
        
        Call .WriteInteger(CharIndex)
        Call .WriteBoolean(invisible)
        
        PrepareMessageSetInvisible = .ReadASCIIStringFixed(.Length)

    End With

End Function

Public Function PrepareMessageCharacterChangeNick(ByVal CharIndex As Integer, _
                                                  ByVal newNick As String) As String

    '***************************************************
    'Author: Budi
    'Last Modification: 07/23/09
    'Prepares the "Change Nick" message and returns it.
    '***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CharacterChangeNick)
        
        Call .WriteInteger(CharIndex)
        Call .WriteASCIIString(newNick)
        
        PrepareMessageCharacterChangeNick = .ReadASCIIStringFixed(.Length)

    End With

End Function

''
' Prepares the "ChatOverHead" message and returns it.
'
' @param    Chat Text to be displayed over the char's head.
' @param    CharIndex The character uppon which the chat will be displayed.
' @param    Color The color to be used when displaying the chat.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The message is written to no outgoing buffer, but only prepared in a single string to be easily sent to several clients.

Public Function PrepareMessageChatOverHead(ByVal Chat As String, _
                                           ByVal CharIndex As Integer, _
                                           ByVal r As Byte, _
                                           ByVal g As Byte, _
                                           ByVal b As Byte, _
                                           Optional ByVal NoConsole As Boolean = False) As String

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "ChatOverHead" message and returns it.
    '***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ChatOverHead)
        Call .WriteASCIIString(Chat)
        Call .WriteInteger(CharIndex)
        Call .WriteBoolean(NoConsole)
        
        Call .WriteByte(r)
        Call .WriteByte(g)
        Call .WriteByte(b)
        
        PrepareMessageChatOverHead = .ReadASCIIStringFixed(.Length)

    End With

End Function

''
' Prepares the "ConsoleMsg" message and returns it.
'
' @param    Chat Text to be displayed over the char's head.
' @param    FontIndex Index of the FONTTYPE structure to use.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageConsoleMsg(ByVal Chat As String, _
                                         ByVal FontIndex As FontTypeNames) As String

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "ConsoleMsg" message and returns it.
    '***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ConsoleMsg)
        Call .WriteASCIIString(Chat)
        Call .WriteByte(FontIndex)
        
        PrepareMessageConsoleMsg = .ReadASCIIStringFixed(.Length)

    End With

End Function

Public Function PrepareMessageScreenMsg(ByVal msgPrimario As String, _
                                         ByVal msgSecundario As String) As String

    '***************************************************
    'Author: Lorwik
    'Last Modification: 24/04/2021
    'Prepares the "RenderMsg" message and returns it.
    '***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ScreenMsg)
        Call .WriteASCIIString(msgPrimario)
        Call .WriteASCIIString(msgSecundario)

        PrepareMessageScreenMsg = .ReadASCIIStringFixed(.Length)

    End With

End Function

Public Function PrepareCommerceConsoleMsg(ByRef Chat As String, _
                                          ByVal FontIndex As FontTypeNames) As String

    '***************************************************
    'Author: ZaMa
    'Last Modification: 03/12/2009
    'Prepares the "CommerceConsoleMsg" message and returns it.
    '***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CommerceChat)
        Call .WriteASCIIString(Chat)
        Call .WriteByte(FontIndex)
        
        PrepareCommerceConsoleMsg = .ReadASCIIStringFixed(.Length)

    End With

End Function

''
' Prepares the "CreateFX" message and returns it.
'
' @param    UserIndex User to which the message is intended.
' @param    CharIndex Character upon which the FX will be created.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCreateFX(ByVal CharIndex As Integer, _
                                       ByVal FX As Integer, _
                                       ByVal FXLoops As Integer) As String

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "CreateFX" message and returns it
    '***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CreateFX)
        Call .WriteInteger(CharIndex)
        Call .WriteInteger(FX)
        Call .WriteInteger(FXLoops)
        
        PrepareMessageCreateFX = .ReadASCIIStringFixed(.Length)

    End With

End Function

''
' Prepares the "PlayWave" message and returns it.
'
' @param    wave The wave to be played.
' @param    X The X position in map coordinates from where the sound comes.
' @param    Y The Y position in map coordinates from where the sound comes.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessagePlayWave(ByVal wave As Integer, _
                                       ByVal X As Integer, _
                                       ByVal Y As Integer) As String

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 08/08/07
    'Last Modified by: Rapsodius
    'Added X and Y positions for 3D Sounds
    '***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.PlayWave)
        Call .WriteInteger(wave)
        Call .WriteInteger(X)
        Call .WriteInteger(Y)
        
        PrepareMessagePlayWave = .ReadASCIIStringFixed(.Length)

    End With

End Function

''
' Prepares the "GuildChat" message and returns it.
'
' @param    Chat Text to be displayed over the char's head.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageGuildChat(ByVal Chat As String) As String

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "GuildChat" message and returns it
    '***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.GuildChat)
        Call .WriteASCIIString(Chat)
        
        PrepareMessageGuildChat = .ReadASCIIStringFixed(.Length)

    End With

End Function

''
' Prepares the "ShowMessageBox" message and returns it.
'
' @param    Message Text to be displayed in the message box.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageShowMessageBox(ByVal Chat As String) As String

    '***************************************************
    'Author: Fredy Horacio Treboux (liquid)
    'Last Modification: 01/08/07
    'Prepares the "ShowMessageBox" message and returns it
    '***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ShowMessageBox)
        Call .WriteASCIIString(Chat)
        
        PrepareMessageShowMessageBox = .ReadASCIIStringFixed(.Length)

    End With

End Function

''
' Prepares the "PlayMidi" message and returns it.
'
' @param    midi The midi to be played.
' @param    loops Number of repets for the midi.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessagePlayMusic(ByVal music As Integer, _
                                       Optional ByVal loops As Integer = -1) As String

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "PlayMidi" message and returns it
    '***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.PlayMusic)
        Call .WriteInteger(music)
        Call .WriteInteger(loops)
        
        PrepareMessagePlayMusic = .ReadASCIIStringFixed(.Length)

    End With

End Function

''
' Prepares the "PauseToggle" message and returns it.
'
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessagePauseToggle() As String

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "PauseToggle" message and returns it
    '***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.PauseToggle)
        PrepareMessagePauseToggle = .ReadASCIIStringFixed(.Length)

    End With

End Function

''
' Prepares the "ActualizarClima" message and returns it.
'
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageActualizarClima() As String

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "ActualizarClima" message and returns it
    '***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ActualizarClima)
        Call .WriteByte(DayStatus)
        
        PrepareMessageActualizarClima = .ReadASCIIStringFixed(.Length)

    End With

End Function

''
' Prepares the "ObjectDelete" message and returns it.
'
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageObjectDelete(ByVal X As Integer, ByVal Y As Integer) As String

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "ObjectDelete" message and returns it
    '***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ObjectDelete)
        Call .WriteInteger(X)
        Call .WriteInteger(Y)
        
        PrepareMessageObjectDelete = .ReadASCIIStringFixed(.Length)

    End With

End Function

''
' Prepares the "BlockPosition" message and returns it.
'
' @param    X X coord of the tile to block/unblock.
' @param    Y Y coord of the tile to block/unblock.
' @param    Blocked Blocked status of the tile
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageBlockPosition(ByVal X As Integer, _
                                            ByVal Y As Integer, _
                                            ByVal Blocked As Boolean) As String

    '***************************************************
    'Author: Fredy Horacio Treboux (liquid)
    'Last Modification: 01/08/07
    'Prepares the "BlockPosition" message and returns it
    '***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.BlockPosition)
        Call .WriteInteger(X)
        Call .WriteInteger(Y)
        Call .WriteBoolean(Blocked)
        
        PrepareMessageBlockPosition = .ReadASCIIStringFixed(.Length)

    End With
    
End Function

''
' Prepares the "ObjectCreate" message and returns it.
'
' @param    GrhIndex Grh of the object.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageObjectCreate(ByVal GrhIndex As Long, _
                                           ByVal ParticulaIndex As Integer, _
                                           ByVal X As Integer, _
                                           ByVal Y As Integer, _
                                           ByVal Shadow As Byte) As String

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'prepares the "ObjectCreate" message and returns it
    '***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ObjectCreate)
        Call .WriteInteger(X)
        Call .WriteInteger(Y)
        Call .WriteLong(GrhIndex)
        Call .WriteInteger(ParticulaIndex)
        Call .WriteByte(Shadow)
        
        PrepareMessageObjectCreate = .ReadASCIIStringFixed(.Length)

    End With

End Function

''
' Prepares the "CharacterRemove" message and returns it.
'
' @param    CharIndex Character to be removed.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCharacterRemove(ByVal CharIndex As Integer) As String

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "CharacterRemove" message and returns it
    '***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CharacterRemove)
        Call .WriteInteger(CharIndex)
        
        PrepareMessageCharacterRemove = .ReadASCIIStringFixed(.Length)

    End With

End Function

''
' Prepares the "RemoveCharDialog" message and returns it.
'
' @param    CharIndex Character whose dialog will be removed.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageRemoveCharDialog(ByVal CharIndex As Integer) As String

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RemoveCharDialog" message to the given user's outgoing data buffer
    '***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.RemoveCharDialog)
        Call .WriteInteger(CharIndex)
        
        PrepareMessageRemoveCharDialog = .ReadASCIIStringFixed(.Length)

    End With

End Function

''
' Writes the "CharacterCreate" message to the given user's outgoing data buffer.
'
' @param    body Body index of the new character.
' @param    head Head index of the new character.
' @param    heading Heading in which the new character is looking.
' @param    CharIndex The index of the new character.
' @param    X X coord of the new character's position.
' @param    Y Y coord of the new character's position.
' @param    weapon Weapon index of the new character.
' @param    shield Shield index of the new character.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @param    helmet Helmet index of the new character.
' @param    name Name of the new character.
' @param    NickColor Determines if the character is a criminal or not, and if can be atacked by someone
' @param    privileges Sets if the character is a normal one or any kind of administrative character.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCharacterCreate(ByVal body As Integer, _
                                              ByVal Head As Integer, _
                                              ByVal Heading As eHeading, _
                                              ByVal CharIndex As Integer, _
                                              ByVal X As Integer, _
                                              ByVal Y As Integer, _
                                              ByVal weapon As Integer, _
                                              ByVal shield As Integer, _
                                              ByVal FX As Integer, _
                                              ByVal FXLoops As Integer, _
                                              ByVal helmet As Integer, _
                                              ByVal AnimAtaque As Integer, _
                                              ByVal Name As String, _
                                              ByVal NickColor As Byte, _
                                              ByVal Privileges As Byte, _
                                              ByVal GrhAura As Long, _
                                              ByVal AuraColor As Long, _
                                              ByVal speeding As Single, _
                                              ByVal estadoQuest As Byte) As String

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "CharacterCreate" message and returns it
    '***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CharacterCreate)
        
        Call .WriteInteger(CharIndex)
        Call .WriteInteger(body)
        Call .WriteInteger(Head)
        Call .WriteByte(Heading)
        Call .WriteInteger(X)
        Call .WriteInteger(Y)
        Call .WriteInteger(weapon)
        Call .WriteInteger(shield)
        Call .WriteInteger(helmet)
        Call .WriteInteger(AnimAtaque)
        Call .WriteInteger(FX)
        Call .WriteInteger(FXLoops)
        Call .WriteASCIIString(Name)
        Call .WriteByte(NickColor)
        Call .WriteByte(Privileges)
        Call .WriteLong(GrhAura)
        Call .WriteLong(AuraColor)
        Call .WriteLong(speeding)
        Call .WriteByte(estadoQuest)
        
        PrepareMessageCharacterCreate = .ReadASCIIStringFixed(.Length)

    End With

End Function

''
' Prepares the "CharacterChange" message and returns it.
'
' @param    body Body index of the new character.
' @param    head Head index of the new character.
' @param    heading Heading in which the new character is looking.
' @param    CharIndex The index of the new character.
' @param    weapon Weapon index of the new character.
' @param    shield Shield index of the new character.
' @param    FX FX index to be displayed over the new character.
' @param    FXLoops Number of times the FX should be rendered.
' @param    helmet Helmet index of the new character.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCharacterChange(ByVal body As Integer, _
                                              ByVal Head As Integer, _
                                              ByVal Heading As eHeading, _
                                              ByVal CharIndex As Integer, _
                                              ByVal weapon As Integer, _
                                              ByVal shield As Integer, _
                                              ByVal FX As Integer, _
                                              ByVal FXLoops As Integer, _
                                              ByVal helmet As Integer, _
                                              ByVal AuraAnim As Long, _
                                              ByVal AuraColor As Long, _
                                              Optional ByVal QuestStatus As Byte = 255) As String

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "CharacterChange" message and returns it
    '***************************************************
    
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CharacterChange)
        
        Call .WriteInteger(CharIndex)
        Call .WriteByte(Heading)
        Call .WriteInteger(body)
        Call .WriteInteger(Head)
        Call .WriteInteger(weapon)
        Call .WriteInteger(shield)
        Call .WriteInteger(helmet)
        Call .WriteInteger(FX)
        Call .WriteInteger(FXLoops)
        Call .WriteLong(AuraAnim)
        Call .WriteLong(AuraColor)
        Call .WriteByte(QuestStatus)
        
        PrepareMessageCharacterChange = .ReadASCIIStringFixed(.Length)

    End With

End Function
Public Function PrepareMessageHeadingChange(ByVal Heading As eHeading, _
                                            ByVal CharIndex As Integer)

    '***************************************************
    'Author: FrankoH298
    'Last Modification: 10/09/19
    'Prepares the "HeadingChange" message and returns it
    '***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.HeadingChange)
        Call .WriteInteger(CharIndex)
        Call .WriteByte(Heading)

        PrepareMessageHeadingChange = .ReadASCIIStringFixed(.Length)

    End With

End Function

''
' Prepares the "CharacterMove" message and returns it.
'
' @param    CharIndex Character which is moving.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageCharacterMove(ByVal CharIndex As Integer, _
                                            ByVal X As Integer, _
                                            ByVal Y As Integer) As String

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "CharacterMove" message and returns it
    '***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CharacterMove)
        Call .WriteInteger(CharIndex)
        Call .WriteInteger(X)
        Call .WriteInteger(Y)
        
        PrepareMessageCharacterMove = .ReadASCIIStringFixed(.Length)

    End With

End Function

Public Function PrepareMessageForceCharMove(ByVal Direccion As eHeading) As String

    '***************************************************
    'Author: ZaMa
    'Last Modification: 26/03/2009
    'Prepares the "ForceCharMove" message and returns it
    '***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.ForceCharMove)
        Call .WriteByte(Direccion)
        
        PrepareMessageForceCharMove = .ReadASCIIStringFixed(.Length)

    End With

End Function

''
' Prepares the "UpdateTagAndStatus" message and returns it.
'
' @param    CharIndex Character which is moving.
' @param    X X coord of the character's new position.
' @param    Y Y coord of the character's new position.
' @return   The formated message ready to be writen as is on outgoing buffers.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageUpdateTagAndStatus(ByVal UserIndex As Integer, _
                                                 ByVal NickColor As Byte, _
                                                 ByRef Tag As String) As String

    '***************************************************
    'Author: Alejandro Salvo (Salvito)
    'Last Modification: 04/07/07
    'Last Modified By: Juan Martin Sotuyo Dodero (Maraxus)
    'Prepares the "UpdateTagAndStatus" message and returns it
    '15/01/2010: ZaMa - Now sends the nick color instead of the status.
    '***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.UpdateTagAndStatus)
        
        Call .WriteInteger(UserList(UserIndex).Char.CharIndex)
        Call .WriteByte(NickColor)
        Call .WriteASCIIString(Tag)
        
        PrepareMessageUpdateTagAndStatus = .ReadASCIIStringFixed(.Length)

    End With

End Function

''
' Prepares the "ErrorMsg" message and returns it.
'
' @param    message The error message to be displayed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Function PrepareMessageErrorMsg(ByVal Message As String) As String

    '***************************************************
    'Author: Juan Martin Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Prepares the "ErrorMsg" message and returns it
    '***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.errorMsg)
        Call .WriteASCIIString(Message)
        
        PrepareMessageErrorMsg = .ReadASCIIStringFixed(.Length)

    End With

End Function

''
' Writes the "StopWorking" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.

Public Sub WriteStopWorking(ByVal UserIndex As Integer)

    '***************************************************
    'Author: ZaMa
    'Last Modification: 21/02/2010
    '
    '***************************************************
    On Error GoTo errHandler
    
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.StopWorking)
        
    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "CancelOfferItem" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @param    Slot      The slot to cancel.

Public Sub WriteCancelOfferItem(ByVal UserIndex As Integer, ByVal Slot As Byte)

    '***************************************************
    'Author: Torres Patricio (Pato)
    'Last Modification: 05/03/2010
    '
    '***************************************************
    On Error GoTo errHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.CancelOfferItem)
        Call .WriteByte(Slot)

    End With
    
    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "RecordDetails" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRecordDetails(ByVal UserIndex As Integer, ByVal RecordIndex As Integer)

    '***************************************************
    'Author: Amraphen
    'Last Modification: 29/11/2010
    'Writes the "RecordDetails" message to the given user's outgoing data buffer
    '***************************************************
    Dim i        As Long

    Dim tIndex   As Integer

    Dim tmpStr   As String

    Dim TempDate As Date

    On Error GoTo errHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.RecordDetails)
        
        'Creador y motivo
        Call .WriteASCIIString(Records(RecordIndex).Creador)
        Call .WriteASCIIString(Records(RecordIndex).Motivo)
        
        tIndex = NameIndex(Records(RecordIndex).Usuario)
        
        'Status del pj (online?)
        Call .WriteBoolean(tIndex > 0)
        
        'Escribo la IP segUn el estado del personaje
        If tIndex > 0 Then
            'La IP Actual
            tmpStr = UserList(tIndex).IP
        Else 'String nulo
            tmpStr = vbNullString

        End If

        Call .WriteASCIIString(tmpStr)
        
        'Escribo tiempo online segUn el estado del personaje
        If tIndex > 0 Then
            'Tiempo logueado.
            TempDate = Now - UserList(tIndex).LogOnTime
            tmpStr = Hour(TempDate) & ":" & Minute(TempDate) & ":" & Second(TempDate)
        Else
            'Envio string nulo.
            tmpStr = vbNullString

        End If

        Call .WriteASCIIString(tmpStr)

        'Escribo observaciones:
        tmpStr = vbNullString

        If Records(RecordIndex).NumObs Then

            For i = 1 To Records(RecordIndex).NumObs
                tmpStr = tmpStr & Records(RecordIndex).Obs(i).Creador & "> " & Records(RecordIndex).Obs(i).Detalles & vbCrLf
            Next i
            
            tmpStr = Left$(tmpStr, Len(tmpStr) - 1)

        End If

        Call .WriteASCIIString(tmpStr)

    End With

    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Writes the "RecordList" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRecordList(ByVal UserIndex As Integer)

    '***************************************************
    'Author: Amraphen
    'Last Modification: 29/11/2010
    'Writes the "RecordList" message to the given user's outgoing data buffer
    '***************************************************
    Dim i As Long

    On Error GoTo errHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.RecordList)
        
        Call .WriteByte(NumRecords)

        For i = 1 To NumRecords
            Call .WriteASCIIString(Records(i).Usuario)
        Next i

    End With

    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Public Sub WriteEnviarPJUserAccount(ByVal UserIndex As Integer)
'***************************************************
'Author: Juan Andres Dalmasso (CHOTS)
'Last Modification: 12/10/2018
'Writes the "AccountLogged" message to the given user with the data of the account he just logged in
'***************************************************
    On Error GoTo errHandler

    Dim i As Long

    With UserList(UserIndex)
        Call .outgoingData.WriteByte(ServerPacketID.EnviarPJUserAccount)
        .Redundance = RandomNumber(5, 250)
        Call .outgoingData.WriteByte(.Redundance)
        Call .outgoingData.WriteASCIIString(.AccountInfo.username)
        Call .outgoingData.WriteByte(.AccountInfo.NumPjs)
        
        Call .outgoingData.WriteASCIIString(.AccountInfo.VIP)
        Call .outgoingData.WriteBoolean(.AccountInfo.esVIP)

        If .AccountInfo.NumPjs > 0 Then

            For i = 1 To .AccountInfo.NumPjs
                Call .outgoingData.WriteASCIIString(.AccountInfo.AccountPJ(i).Name)
                Call .outgoingData.WriteInteger(.AccountInfo.AccountPJ(i).body)
                Call .outgoingData.WriteInteger(.AccountInfo.AccountPJ(i).Head)
                Call .outgoingData.WriteInteger(.AccountInfo.AccountPJ(i).weapon)
                Call .outgoingData.WriteInteger(.AccountInfo.AccountPJ(i).shield)
                Call .outgoingData.WriteInteger(.AccountInfo.AccountPJ(i).helmet)
                Call .outgoingData.WriteByte(.AccountInfo.AccountPJ(i).Class)
                Call .outgoingData.WriteByte(.AccountInfo.AccountPJ(i).race)
                Call .outgoingData.WriteInteger(.AccountInfo.AccountPJ(i).Map)
                Call .outgoingData.WriteByte(.AccountInfo.AccountPJ(i).level)
                Call .outgoingData.WriteBoolean(.AccountInfo.AccountPJ(i).criminal)
                Call .outgoingData.WriteBoolean(.AccountInfo.AccountPJ(i).dead)
                Call .outgoingData.WriteBoolean(.AccountInfo.AccountPJ(i).gameMaster)
            Next i

        End If
    End With

    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Public Function PrepareMessageCharacterAttackAnim(ByVal CharIndex As Integer) As String

    '***************************************************
    'Author: Cucsijuan
    'Last Modification: 2/9/2018
    'Prepares the Attack animation message.
    '***************************************************
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.PlayAttackAnim)
        Call .WriteInteger(CharIndex)
    
        PrepareMessageCharacterAttackAnim = .ReadASCIIStringFixed(.Length)

    End With

End Function

Public Function PrepareMessageFXtoMap(ByVal FxIndex As Integer, _
                                      ByVal loops As Byte, _
                                      ByVal X As Integer, _
                                      ByVal Y As Integer) As String

    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.FXtoMap)
        Call .WriteByte(loops)
        Call .WriteInteger(X)
        Call .WriteInteger(Y)
        Call .WriteInteger(FxIndex)
        
        PrepareMessageFXtoMap = .ReadASCIIStringFixed(.Length)

    End With

End Function

Public Function WriteSearchList(ByVal UserIndex As Integer, _
                                ByVal Num As Integer, _
                                ByVal Datos As String, _
                                ByVal obj As Boolean) As String
 
    On Error GoTo errHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.SearchList)
        Call .WriteInteger(Num)
        Call .WriteBoolean(obj)
        Call .WriteASCIIString(Datos)

    End With
 
errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)

        Resume

    End If
   
End Function

Public Sub WriteQuestDetails(ByVal UserIndex As Integer, _
                             ByVal QuestIndex As Integer, _
                             Optional Questslot As Byte = 0)

    '*******************************************************
    'Envï¿½a el paquete QuestDetails y la informaciï¿½n correspondiente.
    'Last modified: 30/01/2010 by Amraphen
    '*******************************************************
    Dim i As Integer
 
    On Error GoTo errHandler

    With UserList(UserIndex).outgoingData
        'ID del paquete
        Call .WriteByte(ServerPacketID.QuestDetails)
        
        'Se usa la variable QuestSlot para saber si enviamos la info de una quest ya empezada o la info de una quest que no se aceptï¿½ todavï¿½a (1 para el primer caso y 0 para el segundo)
        Call .WriteByte(IIf(Questslot, 1, 0))

        'Enviamos nombre, descripciï¿½n y nivel requerido de la quest
        Call .WriteASCIIString(QuestList(QuestIndex).Nombre)
        Call .WriteASCIIString(QuestList(QuestIndex).Desc)
        Call .WriteByte(QuestList(QuestIndex).RequiredLevel)
        
        'Enviamos la cantidad de npcs a matar
        Call .WriteByte(QuestList(QuestIndex).RequiredNPCs)

        If QuestList(QuestIndex).RequiredNPCs Then
            'Si hay npcs entonces enviamos la lista
            For i = 1 To QuestList(QuestIndex).RequiredNPCs
                Call .WriteInteger(QuestList(QuestIndex).RequiredNPC(i).Amount)
                Call .WriteASCIIString(GetVar(DatPath & "NPCs.dat", "NPC" & QuestList(QuestIndex).RequiredNPC(i).NPCIndex, "Name"))

                'Si es una quest ya empezada, entonces mandamos los NPCs que matï¿½.
                If Questslot Then
                    Call .WriteInteger(UserList(UserIndex).QuestStats.Quests(UserList(UserIndex).QuestStats.QuestEnCurso(Questslot)).NPCsKilled(i))

                End If

            Next i
        End If
        
        'Enviamos la cantidad de npcs a hablar
        Call .WriteByte(QuestList(QuestIndex).RequiredTargetNPCs)
        
        If QuestList(QuestIndex).RequiredTargetNPCs Then
            'Si hay objetivos entonces enviamos la lista
            For i = 1 To QuestList(QuestIndex).RequiredTargetNPCs
                Call .WriteASCIIString(GetVar(DatPath & "NPCs.dat", "NPC" & QuestList(QuestIndex).RequiredTargetNPC(i).NPCIndex, "Name"))

                'Si es una quest ya empezada, entonces mandamos los NPCs que matï¿½.
                If Questslot Then
                    Call .WriteInteger(UserList(UserIndex).QuestStats.Quests(UserList(UserIndex).QuestStats.QuestEnCurso(Questslot)).NPCsTarget(i))

                End If

            Next i

        End If
        
        'Enviamos la cantidad de objs requeridos
        Call .WriteByte(QuestList(QuestIndex).RequiredOBJs)

        If QuestList(QuestIndex).RequiredOBJs Then

            'Si hay objs entonces enviamos la lista
            For i = 1 To QuestList(QuestIndex).RequiredOBJs
                Call .WriteInteger(QuestList(QuestIndex).RequiredOBJ(i).Amount)
                Call .WriteASCIIString(ObjData(QuestList(QuestIndex).RequiredOBJ(i).ObjIndex).Name)
            Next i

        End If
    
        'Enviamos la recompensa de oro y experiencia.
        Call .WriteLong(QuestList(QuestIndex).RewardGLD)
        Call .WriteLong(QuestList(QuestIndex).RewardEXP)
        
        'Enviamos la cantidad de objs de recompensa
        Call .WriteByte(QuestList(QuestIndex).RewardOBJs)

        If QuestList(QuestIndex).RewardOBJs Then

            'si hay objs entonces enviamos la lista
            For i = 1 To QuestList(QuestIndex).RewardOBJs
                Call .WriteInteger(QuestList(QuestIndex).RewardOBJ(i).Amount)
                Call .WriteASCIIString(ObjData(QuestList(QuestIndex).RewardOBJ(i).ObjIndex).Name)
            Next i

        End If

    End With

    Exit Sub
 
errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub
 
Public Sub WriteQuestListSend(ByVal UserIndex As Integer)

    '*******************************************************
    'Envï¿½a el paquete QuestList y la informaciï¿½n correspondiente.
    'Last modified: 30/01/2010 by Amraphen
    '*******************************************************
    Dim i       As Integer

    Dim tmpStr  As String

    Dim tmpByte As Byte
 
    On Error GoTo errHandler
 
    With UserList(UserIndex)
        .outgoingData.WriteByte ServerPacketID.QuestListSend
        
        For i = 1 To .QuestStats.nQuestCurso

            If .QuestStats.QuestEnCurso(i) > 0 Then
                tmpByte = tmpByte + 1
                tmpStr = tmpStr & QuestList(.QuestStats.Quests(.QuestStats.QuestEnCurso(i)).QuestIndex).Nombre & "-"

            End If

        Next i
        
        'Escribimos la cantidad de quests
        Call .outgoingData.WriteByte(tmpByte)
        
        'Escribimos la lista de quests (sacamos el ï¿½ltimo caracter)
        If tmpByte Then
            Call .outgoingData.WriteASCIIString(Left$(tmpStr, Len(tmpStr) - 1))

        End If

    End With

    Exit Sub
 
errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

Public Sub WriteActualizarNPCQuest(ByVal UserIndex As Integer, ByVal NPCIndex As Integer, ByVal Estado As Byte)
    '*******************************************************
    'Autor: lorwik
    'Fecha: 10/05/2021
    'Descripcion: Manda actualizar el simbolo de quest de los NPC
    '*******************************************************
    On Error GoTo errHandler
 
    With UserList(UserIndex)
        .outgoingData.WriteByte ServerPacketID.ActualizarNPCQuest
        
        Call .outgoingData.WriteInteger(Npclist(NPCIndex).Char.CharIndex)
        Call .outgoingData.WriteByte(Estado)
        
    End With

    Exit Sub
 
errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If
End Sub

Public Function PrepareMessageCreateDamage(ByVal X As Integer, _
                                           ByVal Y As Integer, _
                                           ByVal DamageValue As Long, _
                                           ByVal DamageType As Byte)
 
    ' @ Envia el paquete para crear dano (Y)
 
    With auxiliarBuffer
        .WriteByte ServerPacketID.CreateDamage
        .WriteInteger X
        .WriteInteger Y
        .WriteLong DamageValue
        .WriteByte DamageType
     
        PrepareMessageCreateDamage = .ReadASCIIStringFixed(.Length)
     
    End With
 
End Function

Public Function PrepareMessageSpeeding(ByVal CharIndex As Integer, _
                                       ByVal speeding As Single)
 
    ' @ Envia el paquete para cambiar la velocidad
 
    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.SpeedToChar)
        Call .WriteInteger(CharIndex)
        Call .WriteSingle(speeding)
     
        PrepareMessageSpeeding = .ReadASCIIStringFixed(.Length)
     
    End With
 
End Function

Public Sub WriteUserInEvent(ByVal UserIndex As Integer)
    On Error GoTo errHandler
    
    Call UserList(UserIndex).outgoingData.WriteByte(ServerPacketID.UserInEvent)
    Exit Sub

errHandler:
        If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
            Call FlushBuffer(UserIndex)
            Resume
        End If
End Sub

''
' Writes the "EquitandoToggle" message to the given user's outgoing data buffer.
'
' @param    UserIndex User to which the message is intended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteEquitandoToggle(ByVal UserIndex As Integer)
'***************************************************
'Author: Lorwik
'Last Modification: 23/08/11
'Writes the "EquitandoToggle" message to the given user's outgoing data buffer
'***************************************************
On Error GoTo errHandler
    With UserList(UserIndex)
        Call .outgoingData.WriteByte(ServerPacketID.EquitandoToggle)
        
    End With

    Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteSetSpeed(ByVal UserIndex As Integer)
'***************************************************
'Author: Lorwik
'Last Modification: 23/10/11
'Writes the "EquitandoToggle" message to the given user's outgoing data buffer
'***************************************************

On Error GoTo errHandler


    With UserList(UserIndex)
        Call .outgoingData.WriteByte(ServerPacketID.SetSpeed)
        Call .outgoingData.WriteSingle(.flags.Velocidad)
        
    End With

    Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteCargarListaDeAmigos(ByVal UserIndex As Integer, ByVal Slot As Byte)

    On Error GoTo errHandler

    Dim i As Integer

    With UserList(UserIndex).outgoingData
        
                Call .WriteByte(ServerPacketID.EnviarListDeAmigos)
        Call .WriteByte(Slot)
        Call .WriteASCIIString(UserList(UserIndex).Amigos(Slot).Nombre)

    End With

    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If

End Sub

Public Function PrepareMessageProyectil(ByVal UserIndex As Integer, ByVal CharSending As Integer, ByVal CharRecieved As Integer, ByVal GrhIndex As Long) As String
'*************************************
'Autor: Lorwik
'Fecha:16/05/2020
'*************************************

    With auxiliarBuffer
        .WriteByte (ServerPacketID.proyectil)
        .WriteInteger (CharSending)
        .WriteInteger (CharRecieved)
        .WriteLong (GrhIndex)
        
        PrepareMessageProyectil = .ReadASCIIStringFixed(.Length)
    End With

End Function

Public Sub WriteIniciarSubastasOrConsulta(ByVal UserIndex As Integer)
    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.IniciarSubastaConsulta)
    End With
End Sub

Public Sub WriteConfirmarInstruccion(ByVal UserIndex As Integer, ByVal Mensaje As String)
'***************************************************
'Autor: Lorwik
'Fecha: 19/08/2020
'Descripcion: Pide confirmacion para aprender u olvidar una profesion
'***************************************************

    With UserList(UserIndex).outgoingData
    
        Call .WriteByte(ServerPacketID.ConfirmarInstruccion)
        
        .WriteASCIIString (Mensaje)
        
    End With
End Sub

Public Sub WriteAtaqueNPC(ByVal UserIndex As Integer, ByVal NPCIndex As Integer)
On Error GoTo errHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.AtaqueNPC)
        Call .WriteInteger(NPCIndex)
    End With
    
Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteBattlegrounds(ByVal UserIndex As Integer, ByVal BGs As Boolean)
'**************************************
'Autor Lorwik
'Fecha: 02/05/2022
'Descripción: Envia la variable Battlegrounds al cliente
'**************************************
On Error GoTo errHandler

    With UserList(UserIndex).outgoingData
    
        Call .WriteByte(ServerPacketID.BattleGs)
        Call .WriteBoolean(BGs)
    
    End With

Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
End Sub

Public Sub WriteMostrarShop(ByVal UserIndex As Integer)
'**************************************
'Autor Lorwik
'Fecha: 16/05/2022
'Descripción: Envia la Shop al usuario
'**************************************

On Error GoTo errHandler

    Dim i As Integer

    With UserList(UserIndex).outgoingData
    
        Call .WriteByte(ServerPacketID.MostrarShop)
        Call .WriteInteger(UserList(UserIndex).AccountInfo.Gemas)
        Call .WriteInteger(NUMSHOPS)
        
        For i = 1 To NUMSHOPS
        
            Call .WriteInteger(ObjData(ShopObject(i).ObjIndex).GrhIndex)
            Call .WriteASCIIString(ObjData(ShopObject(i).ObjIndex).Name)
            Call .WriteInteger(ShopObject(i).Amount)
            Call .WriteInteger(ShopObject(i).Valor)
        
        Next i
    
    End With

Exit Sub

errHandler:
    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume
    End If
    
    
End Sub

Public Sub WriteActualizarGemasShop(ByVal UserIndex As Integer)
'****************************************
'Autor: Lorwik
'Fecha: 16/05/2022
'Descripción: Actualiza la cantidad de gemas en la Shop
'****************************************

    With UserList(UserIndex).outgoingData
    
        Call .WriteByte(ServerPacketID.ActualizarGemasShop)
        Call .WriteLong(UserList(UserIndex).AccountInfo.Gemas)
        
    End With

End Sub

Public Sub WriteMostrarPVP(ByVal UserIndex As Integer)
    '****************************************************
    'Autor: Lorwik
    'Fecha: 22/05/2022
    'Descripcion: Envia los datos de PVP del usuario al cliente
    '****************************************************
    
    With UserList(UserIndex)
    
        Call .outgoingData.WriteByte(ServerPacketID.MostrarPVP)
        
        Call .outgoingData.WriteByte(.Stats.ELVPVP)
        Call .outgoingData.WriteInteger(.Stats.ExpPVP)
        Call .outgoingData.WriteInteger(.Stats.ELUPVP)
        Call .outgoingData.WriteLong(.Stats.ELO)
    
    End With
    
End Sub

Public Sub WriteMultiMessage(ByVal UserIndex As Integer, _
                             ByVal MessageIndex As Integer, _
                             Optional ByVal Arg1 As Long, _
                             Optional ByVal Arg2 As Long, _
                             Optional ByVal Arg3 As Long, _
                             Optional ByVal StringArg1 As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo errHandler

    With UserList(UserIndex).outgoingData
        Call .WriteByte(ServerPacketID.MultiMessage)
        Call .WriteByte(MessageIndex)
        
        Select Case MessageIndex

            Case eMessages.NPCSwing, eMessages.NPCKillUser, eMessages.BlockedWithShieldUser, eMessages.BlockedWithShieldother, eMessages.UserSwing, eMessages.SafeModeOn, eMessages.SafeModeOff, eMessages.CombatSafeOff, eMessages.CombatSafeOn, eMessages.NobilityLost, eMessages.CantUseWhileMeditating, eMessages.FinishHome
            
            Case eMessages.NPCHitUser
                Call .WriteByte(Arg1) 'Target
                Call .WriteInteger(Arg2) 'damage
                
            Case eMessages.UserHitNPC
                Call .WriteLong(Arg1) 'damage
                
            Case eMessages.UserAttackedSwing
                Call .WriteInteger(UserList(Arg1).Char.CharIndex)
                
            Case eMessages.UserHittedByUser
                Call .WriteInteger(Arg1) 'AttackerIndex
                Call .WriteByte(Arg2) 'Target
                Call .WriteInteger(Arg3) 'damage
                
            Case eMessages.UserHittedUser
                Call .WriteInteger(Arg1) 'AttackerIndex
                Call .WriteByte(Arg2) 'Target
                Call .WriteInteger(Arg3) 'damage
                
            Case eMessages.WorkRequestTarget
                Call .WriteByte(Arg1) 'skill
            
            Case eMessages.HaveKilledUser '"Has matado a " & UserList(VictimIndex).name & "!" "Has ganado " & DaExp & " puntos de experiencia."
                Call .WriteInteger(UserList(Arg1).Char.CharIndex) 'VictimIndex
                Call .WriteLong(Arg2) 'Expe
            
            Case eMessages.UserKill '"" & .name & " te ha matado!"
                Call .WriteInteger(UserList(Arg1).Char.CharIndex) 'AttackerIndex
            
            Case eMessages.EarnExp
            
            Case eMessages.Home
                Call .WriteByte(CByte(Arg1))
                Call .WriteInteger(CInt(Arg2))
                'El cliente no conoce nada sobre nombre de mapas y hogares, por lo tanto _
                 hasta que no se pasen los dats e .INFs al cliente, esto queda asi.
                Call .WriteASCIIString(StringArg1) 'Call .WriteByte(CByte(Arg2))
                
            Case eMessages.UserMuerto
            
            Case eMessages.NpcInmune
            
            Case eMessages.Hechizo_HechiceroMSG_NOMBRE
                Call .WriteByte(CByte(Arg1)) 'SpellIndex
                Call .WriteASCIIString(StringArg1) 'Persona
             
            Case eMessages.Hechizo_HechiceroMSG_ALGUIEN
                Call .WriteByte(CByte(Arg1)) 'SpellIndex
             
            Case eMessages.Hechizo_HechiceroMSG_CRIATURA
                Call .WriteByte(CByte(Arg1)) 'SpellIndex
             
            Case eMessages.Hechizo_PropioMSG
                Call .WriteByte(CByte(Arg1)) 'SpellIndex
         
            Case eMessages.Hechizo_TargetMSG
                Call .WriteByte(CByte(Arg1)) 'SpellIndex
                Call .WriteASCIIString(StringArg1) 'Persona

        End Select

    End With

    Exit Sub ''

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If

End Sub

''
' Prepares the "CreateParticleChar" message and returns it.

Public Function PrepareMessageCreateParticleChar(ByVal CharIndex As Integer, _
                                       ByVal ParticulaID As Integer, _
                                       ByVal Create As Boolean, _
                                       ByVal Life As Long) As String

    '***********************************
    'Autor: Lorwik
    'Fecha: 20/07/2020
    'Descripcion: Enviamos crear una particula en un char
    '***********************************

    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.CharParticle)
        Call .WriteInteger(ParticulaID)
        Call .WriteBoolean(Create)
        Call .WriteInteger(CharIndex)
        Call .WriteLong(Life)
        
        PrepareMessageCreateParticleChar = .ReadASCIIStringFixed(.Length)

    End With

End Function

Public Function PrepareMessageBarFx(ByVal CharIndex As Integer, _
                                    ByVal BarTime As Integer, _
                                    ByVal BarAccion As Byte)

    With auxiliarBuffer
        Call .WriteByte(ServerPacketID.eBarFx)
        Call .WriteInteger(CharIndex)
        Call .WriteInteger(BarTime)
        Call .WriteInteger(BarAccion)
        
        PrepareMessageBarFx = .ReadASCIIStringFixed(.Length)
        
    End With

End Function

Public Sub WritePrivilegios(ByVal UserIndex As Integer)
    '****************************************************
    'Autor: Lorwik
    'Fecha: 28/10/2023
    'Descripcion: Envia si es usuario o gm
    '****************************************************
    On Error GoTo errHandler

    With UserList(UserIndex)

        Call .outgoingData.WriteByte(ServerPacketID.ePrivilegios)
        
        If .flags.Privilegios And (PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero) Then
            Call .outgoingData.WriteBoolean(True)
        Else
            Call .outgoingData.WriteBoolean(False)
        End If
    
    End With
    
    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If
    
End Sub

Public Sub WriteDeletedChar(ByVal UserIndex As Integer)
    '****************************************************
    'Autor: Lorwik
    'Fecha: 28/10/2023
    '****************************************************
    On Error GoTo errHandler

    With UserList(UserIndex)
        Call .outgoingData.WriteByte(ServerPacketID.DeletedChar)
    
    End With

    Exit Sub

errHandler:

    If Err.Number = UserList(UserIndex).outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer(UserIndex)
        Resume

    End If
End Sub
