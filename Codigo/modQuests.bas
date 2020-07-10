Attribute VB_Name = "Quests"

'¿Como funciona este sistema de Quest?
'Cada usuario tiene X cantidad de slots de quest que esta enumerado desde 1 a MAXQUESTS.
'Cada quest se guardara en el slot correspondiente segun su indice. Por ejemplo, una quest
'con ID 5, se guardara en el slot 5.

Option Explicit
 
Public Enum eStatusQuest

    NoAceptada = 0
    EnCurso = 1
    Terminada = 2
        
End Enum
 
'Constantes de las quests
Public Const MAXUSERQUESTS As Integer = 5      'Maxima cantidad de quests aceptadas sin completar que puede tener un usuario al mismo tiempo.
Public Const MAXQUESTS As Integer = 200        'Maxima cantidad de quests que puede tener un usuario
Public NumQuests As Integer                    'Num de quest dateadas actualmente
 
Public Function TieneQuest(ByVal UserIndex As Integer, _
                           ByVal QuestNumber As Integer) As Byte

    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Devuelve el slot de UserQuests en que tiene la quest QuestNumber. En caso contrario devuelve 0.
    'Last modified: 27/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    
    If UserList(UserIndex).QuestStats.Quests(QuestNumber).QuestStatus = eStatusQuest.EnCurso Then
        TieneQuest = QuestNumber
        Exit Function

    End If
    
    TieneQuest = 0

End Function

Public Function EstadoQuest(ByVal UserIndex As Integer, _
                           ByVal QuestNumber As Integer) As Byte

    '****************************************
    'Autor: Lorwik
    'Fecha: 09/07/2020
    'Descripcion: Devuelve el estado de la quest
    '****************************************
    
    'Si el numero de la quest es invalida, devolvemos como que no acepto ninguna quest
    If QuestNumber <= 0 Or UserIndex = 0 Then
        EstadoQuest = 0
        Exit Function
    End If
    
    EstadoQuest = UserList(UserIndex).QuestStats.Quests(QuestNumber).QuestStatus
    
End Function

Private Function MaxQuestsAceptadas(ByVal UserIndex As Integer) As Boolean
    '****************************************
    'Autor: Lorwik
    'Fecha: 28/06/2020
    'Descripcion: Comprueba si el usuario supero el numero maximo de quest en curso
    '****************************************

    Dim i As Integer
    Dim Count As Integer
    
    For i = 1 To MAXQUESTS
        If UserList(UserIndex).QuestStats.Quests(i).QuestStatus = eStatusQuest.EnCurso Then Count = Count + 1
    Next i
    
    If Count >= MAXUSERQUESTS Then
        MaxQuestsAceptadas = True
    Else
        MaxQuestsAceptadas = False
    End If

End Function
 
Public Sub HandleQuestAccept(ByVal UserIndex As Integer)

    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Maneja el evento de aceptar una quest.
    'Last modified: 31/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    Dim NpcIndex  As Integer

    Dim QuestSlot As Byte
    
    Dim i         As Byte
 
    Call UserList(UserIndex).incomingData.ReadByte
 
    NpcIndex = UserList(UserIndex).flags.TargetNPC
    
    If NpcIndex = 0 Then Exit Sub
    
    'Esta el personaje en la distancia correcta?
    If Distancia(UserList(UserIndex).Pos, Npclist(NpcIndex).Pos) > 5 Then
        Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
'
    End If
    
    'Esta el user muerto?
    If UserList(UserIndex).flags.Muerto = 1 Then
        Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
        Exit Sub

    End If
    
    'Agregamos la quest.
    With UserList(UserIndex).QuestStats.Quests(Npclist(NpcIndex).QuestNumber)
        .QuestStatus = eStatusQuest.EnCurso
        
        If QuestList(Npclist(NpcIndex).QuestNumber).RequiredNPCs Then
            ReDim .NPCsKilled(1 To QuestList(Npclist(NpcIndex).QuestNumber).RequiredNPCs)
            
            For i = 1 To QuestList(Npclist(NpcIndex).QuestNumber).RequiredNPCs
                .NPCsKilled(i) = 0
            Next i
            
        End If
        
        Call WriteConsoleMsg(UserIndex, "Has aceptado la mision " & Chr(34) & QuestList(Npclist(NpcIndex).QuestNumber).Nombre & Chr(34) & ".", FontTypeNames.FONTTYPE_INFO)
        
    End With

End Sub
 
Public Sub FinishQuest(ByVal UserIndex As Integer, _
                       ByVal QuestIndex As Integer, _
                       ByVal QuestSlot As Byte)

    '****************************************************
    'Autor: Amraphen
    'Fecha: 29/01/2010
    'Maneja el evento de terminar una quest.
    '****************************************************
    
    Dim i              As Integer

    Dim InvSlotsLibres As Byte

    Dim NpcIndex       As Integer
 
    NpcIndex = UserList(UserIndex).flags.TargetNPC
    
    With QuestList(QuestIndex)

        'Esta el user muerto?
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
            Exit Sub
    
        End If
    
        'Comprobamos que tenga los objetos.
        If .RequiredOBJs > 0 Then

            For i = 1 To .RequiredOBJs

                If TieneObjetos(.RequiredOBJ(i).ObjIndex, .RequiredOBJ(i).Amount, UserIndex) = False Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("No has conseguido todos los objetos que te he pedido.", Npclist(NpcIndex).Char.CharIndex, vbWhite))
                    Exit Sub

                End If

            Next i

        End If
        
        'Comprobamos que haya matado todas las criaturas.
        If .RequiredNPCs > 0 Then

            For i = 1 To .RequiredNPCs

                If .RequiredNPC(i).Amount > UserList(UserIndex).QuestStats.Quests(QuestSlot).NPCsKilled(i) Then
                    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("No has matado todas las criaturas que te he pedido.", Npclist(NpcIndex).Char.CharIndex, vbWhite))
                    Exit Sub

                End If

            Next i

        End If
    
        'Comprobamos que el usuario tenga espacio para recibir los items.
        If .RewardOBJs > 0 Then

            'Buscamos la cantidad de slots de inventario libres.
            For i = 1 To MAX_INVENTORY_SLOTS

                If UserList(UserIndex).Invent.Object(i).ObjIndex = 0 Then InvSlotsLibres = InvSlotsLibres + 1
            Next i
            
            'Nos fijamos si entra
            If InvSlotsLibres < .RewardOBJs Then
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("No tienes suficiente espacio en el inventario para recibir la recompensa. Vuelve cuando hayas hecho mas espacio.", Npclist(NpcIndex).Char.CharIndex, vbWhite))
                Exit Sub

            End If

        End If
    
        'A esta altura ya cumplio los objetivos, entonces se le entregan las recompensas.
        Call WriteConsoleMsg(UserIndex, "Has completado la mision " & Chr(34) & QuestList(QuestIndex).Nombre & Chr(34) & "!", FontTypeNames.FONTTYPE_INFO)
        
        'Si la quest pedia objetos, se los saca al personaje.
        If .RequiredOBJs Then

            For i = 1 To .RequiredOBJs
                Call QuitarObjetos(.RequiredOBJ(i).ObjIndex, .RequiredOBJ(i).Amount, UserIndex)
            Next i

        End If
        
        'Se entrega la experiencia.
        If .RewardEXP Then
            UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + .RewardEXP
            Call WriteConsoleMsg(UserIndex, "Has ganado " & .RewardEXP & " puntos de experiencia como recompensa.", FontTypeNames.FONTTYPE_INFO)

        End If
        
        'Se entrega el oro.
        If .RewardGLD Then
            UserList(UserIndex).Stats.Gld = UserList(UserIndex).Stats.Gld + .RewardGLD
            Call WriteConsoleMsg(UserIndex, "Has ganado " & .RewardGLD & " monedas de oro como recompensa.", FontTypeNames.FONTTYPE_INFO)

        End If
        
        'Si hay recompensa de objetos, se entregan.
        If .RewardOBJs > 0 Then

            For i = 1 To .RewardOBJs

                If .RewardOBJ(i).Amount Then
                    Call MeterItemEnInventario(UserIndex, .RewardOBJ(i))
                    Call WriteConsoleMsg(UserIndex, "Has recibido " & QuestList(QuestIndex).RewardOBJ(i).Amount & " " & ObjData(QuestList(QuestIndex).RewardOBJ(i).ObjIndex).Name & " como recompensa.", FontTypeNames.FONTTYPE_INFO)

                End If

            Next i

        End If

        'Se agrega que el usuario ya hizo esta quest.
        UserList(UserIndex).QuestStats.Quests(QuestIndex).QuestStatus = eStatusQuest.Terminada
        UserList(UserIndex).QuestStats.NumQuestsDone = UserList(UserIndex).QuestStats.NumQuestsDone + 1

        'Actualizamos el personaje
        Call CheckUserLevel(UserIndex)
        Call UpdateUserInv(True, UserIndex, 0)

    End With

End Sub
 
Public Function UserDoneQuest(ByVal UserIndex As Integer, _
                              ByVal QuestIndex As Integer) As Boolean

    '****************************************************
    'Autor: Lorwik
    'Fecha: 28/06/2020
    'Descripcion: Verifica si el usuario hizo la quest QuestIndex.
    '****************************************************
    With UserList(UserIndex).QuestStats
    
        'Tiene la quest terminada?
        If .Quests(QuestIndex).QuestStatus = eStatusQuest.Terminada Then
            UserDoneQuest = True
            Exit Function

        End If

    End With
    
    UserDoneQuest = False
        
End Function
 
Public Sub CleanQuestSlot(ByVal UserIndex As Integer, ByVal QuestSlot As Integer)

    '****************************************************
    'Autor: Lorwik
    'Fecha: 28/06/2020
    'Descripcion: Limpia un slot de quest de un usuario.
    '****************************************************
    Dim i As Integer
 
    With UserList(UserIndex).QuestStats.Quests(QuestSlot)
        
        If QuestSlot = NumQuests Then Exit Sub
        
        If QuestList(QuestSlot).RequiredNPCs Then

            For i = 1 To QuestList(QuestSlot).RequiredNPCs
                .NPCsKilled(i) = 0
                .QuestStatus = eStatusQuest.NoAceptada
            Next i

        End If
        
    End With

End Sub
 
Public Sub ResetQuestStats(ByVal UserIndex As Integer)

    '****************************************************
    'Autor: Lorwik
    'Fecha: 28/06/2020
    'Limpia todos los QuestStats de un usuario
    '****************************************************
    
    Dim i As Integer
 
    For i = 1 To MAXQUESTS
        Call CleanQuestSlot(UserIndex, i)
    Next i
    
    With UserList(UserIndex).QuestStats
        .NumQuestsDone = 0
    End With

End Sub
 
Public Sub HandleQuest(ByVal UserIndex As Integer)

    '****************************************************
    'Maneja el paquete Quest.
    'Last modified: 18/05/2020
    'Lorwik: Paso todo el chequeo y la accion a otro sub refractorio
    '****************************************************
    
    Dim NpcIndex As Integer

    'Leemos el paquete
    Call UserList(UserIndex).incomingData.ReadByte
 
    NpcIndex = UserList(UserIndex).flags.TargetNPC
    
    Call AccionParaQuest(UserIndex, NpcIndex)

End Sub

Public Sub AccionParaQuest(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
    '****************************************
    'Autor: Lorwik
    'Fecha: 18/05/2020
    'Descripción: 'Refactorizo para que los NPC de quest den misiones
    '****************************************
    
    Dim tmpByte  As Byte

    If NpcIndex = 0 Then Exit Sub
    
    With Npclist(NpcIndex)
    
        'Esta el user muerto?
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
            Exit Sub
    
        End If
    
        'Esta el personaje en la distancia correcta?
        If Distancia(UserList(UserIndex).Pos, .Pos) > 5 Then
            Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
    
        End If
        
        'El NPC hace quests?
        If .QuestNumber = 0 Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("No tengo ninguna mision para ti.", .Char.CharIndex, vbWhite))
            Exit Sub
    
        End If
        
        'El personaje ya hizo la quest?
        If UserDoneQuest(UserIndex, .QuestNumber) Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("Gracias por la ayuda, quizas en otro momento podamos te necesite.", .Char.CharIndex, vbWhite))
            Exit Sub
    
        End If
     
        'El personaje tiene suficiente nivel?
        If UserList(UserIndex).Stats.ELV < QuestList(.QuestNumber).RequiredLevel Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("Debes ser por lo menos nivel " & QuestList(.QuestNumber).RequiredLevel & " para emprender esta mision.", .Char.CharIndex, vbWhite))
            Exit Sub
    
        End If
        
        'A esta altura ya analizo todas las restricciones y esta preparado para el handle propiamente dicho
     
        tmpByte = TieneQuest(UserIndex, .QuestNumber)
        
        If tmpByte Then
            'El usuario esta haciendo la quest, entonces va a hablar con el NPC para recibir la recompensa.
            Call FinishQuest(UserIndex, .QuestNumber, tmpByte)
        Else
            
            'El personaje tiene algun slot de quest para la nueva quest?
            If MaxQuestsAceptadas(UserIndex) Then
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageChatOverHead("Estas haciendo demasiadas misiones. Vuelve cuando hayas completado alguna.", .Char.CharIndex, vbWhite))
                Exit Sub
    
            End If
            
            'Enviamos los detalles de la quest
            Call WriteQuestDetails(UserIndex, .QuestNumber)
    
        End If
    End With
End Sub
 
Public Sub LoadQuests()

    '****************************************************
    'Autor: Amraphen
    'Fecha: 27/01/2010
    'Carga el archivo QUESTS.DAT en el array QuestList.
    '****************************************************
     
    On Error GoTo ErrorHandler

    Dim Reader    As clsIniManager

    Dim tmpStr    As String

    Dim i         As Integer

    Dim j         As Integer
    
    'Cargamos el clsIniManager en memoria
    Set Reader = New clsIniManager
    
    'Lo inicializamos para el archivo Quests.DAT
    Call Reader.Initialize(DatPath & "Quests.DAT")
    
    'Redimensionamos el array
    NumQuests = Reader.GetValue("INIT", "NumQuests")
    ReDim QuestList(1 To NumQuests)

    'Cargamos los datos
    For i = 1 To NumQuests

        With QuestList(i)
            .Nombre = Reader.GetValue("QUEST" & i, "Nombre")
            .Desc = Reader.GetValue("QUEST" & i, "Desc")
            .RequiredLevel = val(Reader.GetValue("QUEST" & i, "RequiredLevel"))
            
            'CARGAMOS OBJETOS REQUERIDOS
            .RequiredOBJs = val(Reader.GetValue("QUEST" & i, "RequiredOBJs"))

            If .RequiredOBJs > 0 Then
                ReDim .RequiredOBJ(1 To .RequiredOBJs)

                For j = 1 To .RequiredOBJs
                    tmpStr = Reader.GetValue("QUEST" & i, "RequiredOBJ" & j)
                    
                    .RequiredOBJ(j).ObjIndex = val(ReadField(1, tmpStr, 45))
                    .RequiredOBJ(j).Amount = val(ReadField(2, tmpStr, 45))
                Next j

            End If
            
            'CARGAMOS NPCS REQUERIDOS
            .RequiredNPCs = val(Reader.GetValue("QUEST" & i, "RequiredNPCs"))

            If .RequiredNPCs > 0 Then
                ReDim .RequiredNPC(1 To .RequiredNPCs)

                For j = 1 To .RequiredNPCs
                    tmpStr = Reader.GetValue("QUEST" & i, "RequiredNPC" & j)
                    
                    .RequiredNPC(j).NpcIndex = val(ReadField(1, tmpStr, 45))
                    .RequiredNPC(j).Amount = val(ReadField(2, tmpStr, 45))
                Next j

            End If
            
            .RewardGLD = val(Reader.GetValue("QUEST" & i, "RewardGLD"))
            .RewardEXP = val(Reader.GetValue("QUEST" & i, "RewardEXP"))
            
            'CARGAMOS OBJETOS DE RECOMPENSA
            .RewardOBJs = val(Reader.GetValue("QUEST" & i, "RewardOBJs"))

            If .RewardOBJs > 0 Then
                ReDim .RewardOBJ(1 To .RewardOBJs)

                For j = 1 To .RewardOBJs
                    tmpStr = Reader.GetValue("QUEST" & i, "RewardOBJ" & j)
                    
                    .RewardOBJ(j).ObjIndex = val(ReadField(1, tmpStr, 45))
                    .RewardOBJ(j).Amount = val(ReadField(2, tmpStr, 45))
                Next j

            End If

        End With

    Next i
    
    'Eliminamos la clase
    Set Reader = Nothing
    Exit Sub
                    
ErrorHandler:
    MsgBox "Error cargando el archivo QUESTS.DAT.", vbOKOnly + vbCritical

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
    
    Dim QuestSlot As Byte
 
    'Leemos el paquete
    Call UserList(UserIndex).incomingData.ReadByte
    
    QuestSlot = UserList(UserIndex).incomingData.ReadByte
    
    Call WriteQuestDetails(UserIndex, QuestSlot, QuestSlot)

End Sub
 
Public Sub HandleQuestAbandon(ByVal UserIndex As Integer)
    '****************************************************
    'Autor: Amraphen
    'Fecha: 31/01/2010
    'Descripcion: Maneja el paquete QuestAbandon.
    '****************************************************
    
    'Leemos el paquete.
    Call UserList(UserIndex).incomingData.ReadByte
    
    'Borramos la quest.
    Call CleanQuestSlot(UserIndex, UserList(UserIndex).incomingData.ReadByte)
    
    'Enviamos la lista de quests actualizada.
    Call WriteQuestListSend(UserIndex)

End Sub
