Attribute VB_Name = "Cuentas"
Option Explicit

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
                             ByVal Password As String, _
                             ByVal Salt As String)

    '***************************************************
    'Autor: Juan Andres Dalmasso (CHOTS)
    'Last Modification: 21/09/2018
    'Saves the password and salt
    '***************************************************

    Call StorePasswordSaltDatabase(UserName, Password, Salt)

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
