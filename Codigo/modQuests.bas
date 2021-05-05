Attribute VB_Name = "modQuests"
' Autor: Lorwik
' Inspirado en el sistema de Quest de Blisse

'¿Como funciona este sistema de Quest?
'Cada usuario tiene X cantidad de slots de quest que esta enumerado desde 1 a MAXQUESTS.
'Cada quest se guardara en el slot correspondiente segun su indice. Por ejemplo, una quest
'con ID 5, se guardara en el slot 5.

'NOTA: Este sistema ha sido documentado hasta la muerte, he refactorizado tanto que ya me hago un lio.

Option Explicit
 
Public Enum eStatusQuest

    NoAceptada = 0
    EnCurso = 1
    Terminada = 2
        
End Enum
 
'Constantes de las quests
Public Const MAXUSERQUESTS As Integer = 5      'Maxima cantidad de quests aceptadas sin completar que puede tener un usuario al mismo tiempo.
Public Const MAXQUESTS As Integer = 250        'Maxima cantidad de quests que puede tener un usuario
Public NumQuests As Integer                    'Num de quest dateadas actualmente
 
Public Sub ListarQuestsenCurso(ByVal UserIndex As Integer)
    
    '****************************************
    'Autor: Lorwik
    'Fecha: 12/08/2020
    'Descripcion: Llena el array de la lista de Quests en curso
    '****************************************
    
    Dim Count As Integer
    Dim i As Integer
    
    With UserList(UserIndex).QuestStats
    
        'Buscamos todas las quests en curso de entre TODAS las quests
        For i = 1 To MAXQUESTS
        
            'Si encuentra una quest en curso, y no supera el limite de quests en curso...
            If .Quests(i).QuestStatus = eStatusQuest.EnCurso And (Count + 1) < MAXUSERQUESTS Then
                .QuestEnCurso(Count + 1) = i
                
                'Aumentamos el contador
                Count = Count + 1
            End If
        
        Next i
    
    End With
    
End Sub
 
Public Sub FinishQuest(ByVal UserIndex As Integer, _
                       ByVal QuestIndex As Integer, _
                       ByVal QuestSlot As Byte)
    '****************************************************
    'Autor: Lorwik
    'Fecha: 12/08/2020
    'Maneja Se comprueba si se dieron todas las condiciones para completar la quest
    'si se dan todas las condiciones se marca como completada y se da la recompensa al usuario.
    '****************************************************
    
    On Error GoTo ErrorHandler
    
    Dim i              As Integer

    Dim InvSlotsLibres As Byte

    Dim NPCIndex       As Integer
 
    NPCIndex = UserList(UserIndex).flags.TargetNPC
    
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
                    Call WriteChatOverHead(UserIndex, "No has conseguido todos los objetos que te he pedido.", Npclist(NPCIndex).Char.CharIndex, vbWhite)
                    Exit Sub

                End If

            Next i

        End If
        
        'Comprobamos que haya matado todas las criaturas.
        If .RequiredNPCs > 0 Then

            For i = 1 To .RequiredNPCs

                If .RequiredNPC(i).Amount > UserList(UserIndex).QuestStats.Quests(QuestSlot).NPCsKilled(i) Then
                    Call WriteChatOverHead(UserIndex, "No has matado todas las criaturas que te he pedido.", Npclist(NPCIndex).Char.CharIndex, vbWhite)
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
                Call WriteChatOverHead(UserIndex, "No tienes suficiente espacio en el inventario para recibir la recompensa. Vuelve cuando hayas hecho mas espacio.", Npclist(NPCIndex).Char.CharIndex, vbWhite)
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
            Call WriteConsoleMsg(UserIndex, "Has ganado " & .RewardEXP & " puntos de experiencia como recompensa.", FontTypeNames.FONTTYPE_EXP)

        End If
        
        'Se entrega el oro.
        If .RewardGLD Then
            UserList(UserIndex).Stats.Gld = UserList(UserIndex).Stats.Gld + .RewardGLD
            Call WriteConsoleMsg(UserIndex, "Has ganado " & .RewardGLD & " monedas de oro como recompensa.", FontTypeNames.FONTTYPE_INFOBOLD)

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
        
        'Eliminamos la quest de la lista de seguimientos
        Call DelQuestEnCurso(UserIndex, QuestIndex)

        'Actualizamos el personaje
        Call CheckUserLevel(UserIndex)
        Call UpdateUserInv(True, UserIndex, 0)
        
        Call ActualizarNPCQuest(UserIndex, NPCIndex)

    End With
    
    Exit Sub
    
ErrorHandler:
    Call LogError("Error accionUseraNPCQuest: " & Err.Number & " - " & Err.description)

End Sub

Private Sub ActualizarNPCQuest(ByVal UserIndex As Integer, ByVal NPCIndex As Integer)

    Dim Estadoquest As Byte

    With Npclist(NPCIndex)
    
        If .QuestNumber(1) > 0 Then
            Estadoquest = modQuests.Estadoquest(UserIndex, .QuestNumber(1))
        Else
            Estadoquest = 255 'El NPC No tiene quest
        End If
    
        Call WriteCharacterChange(UserIndex, .Char.body, .Char.Head, .Char.Heading, .Char.CharIndex, .Char.WeaponAnim, _
            .Char.ShieldAnim, 0, 0, .Char.CascoAnim, 0, 0, Estadoquest)
    
    End With
    
End Sub
 
Public Sub CleanQuestSlot(ByVal UserIndex As Integer, ByVal QuestSlot As Integer)
    '****************************************************
    'Autor: Lorwik
    'Fecha: 28/06/2020
    'Descripcion: Limpia un slot de quest de un usuario.
    '****************************************************
    Dim i As Integer
 
    With UserList(UserIndex).QuestStats
        
        '¿El slot de quest es mayor al numero de quest cargadas o no tiene la quest aceptada?
        If QuestSlot > NumQuests Then Exit Sub
        
        ' Si la quest requiere NPC y esta en la DB del usuario
        If QuestList(QuestSlot).RequiredNPCs And .Quests(QuestSlot).QuestStatus <> eStatusQuest.NoAceptada Then

            For i = 1 To QuestList(QuestSlot).RequiredNPCs
                .Quests(QuestSlot).NPCsKilled(i) = 0
            Next i

        End If
        
        .Quests(QuestSlot).QuestStatus = eStatusQuest.NoAceptada
        
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
    
    For i = 1 To MAXUSERQUESTS
        UserList(UserIndex).QuestStats.QuestEnCurso(i) = 0
    Next i
    
    With UserList(UserIndex).QuestStats
        .NumQuestsDone = 0
    End With

End Sub

Public Sub accionUseraNPCQuest(ByVal UserIndex As Integer, ByVal NPCIndex As Integer)
    '****************************************
    'Autor: Lorwik
    'Fecha: 18/05/2020
    'Descripción: Cuando el usuario interactua con un NPC que tiene quest llega a este sub.
    'Pueden darse 3 situaciones:
    '1- El NPC tiene quest, el usuario no la hizo y la acepta.
    '2- El NPC tiene quest, el usuario la hizo y si el NPC tiene una segunda mision se la dia, si no adios.
    '3- El Usuario acaba de completar la mision que ya le dio el NPC y quiere entregarla.
    '****************************************
    
    On Error GoTo ErrorHandler
    Dim nQuest  As Byte
    Dim SlotLibreQuest As Integer
    Dim NextQuest As Integer
    
    If NPCIndex = 0 Then Exit Sub
    
    With Npclist(NPCIndex)
    
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
        
        'Cadenas de quest:
        NextQuest = BuscarSiguienteQuest(UserIndex, NPCIndex)
        
        'El NPC hace quests?
        If NextQuest = 0 Then
            Call WriteChatOverHead(UserIndex, "No tengo ninguna mision para ti.", .Char.CharIndex, vbWhite)
            Exit Sub
    
        End If
        
        'El personaje ya hizo la quest?
        If UserDoneQuest(UserIndex, NextQuest) Then
            Call WriteChatOverHead(UserIndex, "Gracias por la ayuda, quizas en otro momento podamos te necesite.", .Char.CharIndex, vbWhite)
            Exit Sub
    
        End If
     
        'El personaje tiene suficiente nivel?
        If UserList(UserIndex).Stats.ELV < QuestList(NextQuest).RequiredLevel Then
            Call WriteChatOverHead(UserIndex, "Debes ser por lo menos nivel " & QuestList(NextQuest).RequiredLevel & " para emprender esta mision.", .Char.CharIndex, vbWhite)
            Exit Sub
    
        End If
     
        nQuest = TieneQuest(UserIndex, NextQuest)
     
        '¿Ya tenia la quest, la completo y quiere entregar?
        If nQuest > 1 Then
            'El usuario esta haciendo la quest, entonces va a hablar con el NPC para recibir la recompensa.
            Call FinishQuest(UserIndex, NextQuest, nQuest)
            
        Else '¿No tiene la quest y la quiere aceptar?
            
            'Si obtuvimos -1, es que no hay slots libres
            If MaxQuestsAceptadas(UserIndex) < 0 Then
                Call WriteChatOverHead(UserIndex, "Estas haciendo demasiadas misiones. Vuelve cuando hayas completado alguna.", .Char.CharIndex, vbWhite)
                Exit Sub
    
            End If
            
            'Enviamos los detalles de la quest
            Call WriteQuestDetails(UserIndex, NextQuest)
    
        End If
    End With
    
    Exit Sub
    
ErrorHandler:
    Call LogError("Error accionUseraNPCQuest: " & Err.Number & " - " & Err.description)
    
End Sub

'###########################
'Funciones
'###########################

Public Function BuscarSiguienteQuest(ByVal UserIndex As Integer, ByVal NPCIndex As Integer)
    Dim i As Byte

    'Buscamos la siguiente quest disponible del NPC
    For i = 1 To 5
        If Not UserList(UserIndex).QuestStats.Quests(Npclist(NPCIndex).QuestNumber(i)).QuestStatus = Terminada And Npclist(NPCIndex).QuestNumber(i) <> 0 Then
            BuscarSiguienteQuest = Npclist(NPCIndex).QuestNumber(i)
            Exit Function
        End If
    Next
    
    'No encontro
    BuscarSiguienteQuest = 0
    
End Function

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

Private Function AddQuestEnCurso(ByVal UserIndex As Integer, ByVal Slot As Byte, ByVal NewQuest As Integer) As Boolean
    '****************************************
    'Autor: Lorwik
    'Fecha: 12/08/2020
    'Descripcion: Añade una quest a la lista de seguimientos
    '****************************************
    
    With UserList(UserIndex).QuestStats
    
        If .QuestEnCurso(Slot) = 0 Then
            .QuestEnCurso(Slot) = NewQuest
            AddQuestEnCurso = True
            
        Else
            AddQuestEnCurso = False
            
        End If
    
    End With
End Function

Private Function DelQuestEnCurso(ByVal UserIndex As Integer, ByVal QuestIndex As Integer) As Boolean
    '****************************************
    'Autor: Lorwik
    'Fecha: 15/08/2020
    'Descripcion: Elimina una quest a la lista de seguimientos
    'Recibe el numero de la quest (questindex), y obtiene el slot donde se encuentra
    'esa quest en curso
    '****************************************
    
    Dim i As Integer
    Dim Slot As Byte
    Dim Count As Byte
    Dim tmpArr(1 To MAXUSERQUESTS) As Byte
    
    With UserList(UserIndex).QuestStats
    
        'Buscamos el Slot donde se encuentra la quest
        For i = 1 To MAXUSERQUESTS
            If .QuestEnCurso(i) = QuestIndex Then Slot = i
        Next i
        
        'Si la quest era la ultima en la lista, la eliminamos y listo
        If Slot = MAXUSERQUESTS Then
            .QuestEnCurso(Slot) = 0
            DelQuestEnCurso = True
            Exit Function
        
        Else 'Si no era el ultimo hay que reordenar
            
            Count = 1
            
            For i = 1 To MAXUSERQUESTS
            
                'Vamos copiando al Array temporal en las nuevas posiciones
                If i <> Slot Then
                    tmpArr(Count) = .QuestEnCurso(i)
                    Count = Count + 1
                End If
                
                'Vamos limpiando el Array
                .QuestEnCurso(i) = 0
            Next i
            
            'Volcamos toda la data ordenada en el array de siempre
            For i = 1 To MAXUSERQUESTS
                .QuestEnCurso(i) = tmpArr(i)
            Next i
        
            DelQuestEnCurso = True
            Exit Function
        End If
        
        DelQuestEnCurso = False
    
    End With
    
End Function

Private Function MaxQuestsAceptadas(ByVal UserIndex As Integer) As Integer
    '****************************************
    'Autor: Lorwik
    'Fecha: 28/06/2020
    'Descripcion: Comprueba si el usuario supero el numero maximo de quest en curso
    '****************************************

    Dim i As Integer
    Dim Count As Integer
    
    '¿Tiene hueco en la lista de quests en curso?
    For i = 1 To MAXUSERQUESTS
        If UserList(UserIndex).QuestStats.QuestEnCurso(i) > 0 Then Count = Count + 1
    Next i
    
    'Si no hay hueco, le devolvemos un -1
    If Count >= MAXUSERQUESTS Then
        MaxQuestsAceptadas = -1
        
    Else 'Si hay hueco, devolvemos la posicion libre
        MaxQuestsAceptadas = Count + 1
        
    End If

End Function

Public Function TieneQuest(ByVal UserIndex As Integer, _
                           ByVal nQuest As Integer) As Byte
    '****************************************
    'Autor: Lorwik
    'Fecha: 09/07/2020
    'Descripcion: Devuelve el slot de la quest si esta en curso
    'NOTA: Tambien se podria mirar en la lista de quests en curso, pero esto es valido igual.
    '****************************************
    
    If UserList(UserIndex).QuestStats.Quests(nQuest).QuestStatus = eStatusQuest.EnCurso Then
        TieneQuest = nQuest
        Exit Function

    End If
    
    TieneQuest = 0

End Function

Public Function Estadoquest(ByVal UserIndex As Integer, _
                           ByVal nQuest As Integer) As Byte
    '****************************************
    'Autor: Lorwik
    'Fecha: 09/07/2020
    'Descripcion: Devuelve el estado de la quest
    '****************************************
    
    'Si el numero de la quest es invalida, devolvemos como que no acepto ninguna quest
    If nQuest <= 0 Or UserIndex = 0 Then
        Estadoquest = 0
        Exit Function
    End If
    
    Estadoquest = UserList(UserIndex).QuestStats.Quests(nQuest).QuestStatus
    
End Function

'=================================================
'Funciones de Acciones del usuario sobre quests
'(Acepta, abandona, completa)
'=================================================

Public Function userAbandonaQuest(ByVal UserIndex As Integer, ByVal QuestSlot As Integer, ByVal QuestIndex As Integer) As Boolean
    '****************************************
    'Autor: Lorwik
    'Fecha: 04/05/2021
    'Descripcion: El usuario quiere abandonar la quest
    '****************************************
    
    On Error GoTo ErrorHandler
    
    'Elimiamos la quest de la lista de seguimiento
    Call DelQuestEnCurso(UserIndex, QuestIndex)
        
    Call WriteConsoleMsg(UserIndex, "Has cancelado la mision " & Chr(34) & QuestList(QuestIndex).Nombre & Chr(34) & ".", FontTypeNames.FONTTYPE_INFO)
        
    'Borramos la quest, de la lista global
    Call CleanQuestSlot(UserIndex, QuestIndex)
        
    'Enviamos la lista de quests actualizada.
    Call WriteQuestListSend(UserIndex)
    
    userAbandonaQuest = True
    
    Exit Function
    
ErrorHandler:
    userAbandonaQuest = False
    Call LogError("Error userAbandonaQuest: " & Err.Number & " - " & Err.description)
    
End Function

Public Function userAceptaquest(ByVal UserIndex As Integer) As Boolean
    '****************************************
    'Autor: Lorwik
    'Fecha: 04/05/2021
    'Descripcion: El usuario acepta una nueva quest
    '****************************************

    On Error GoTo ErrorHandler

    Dim NPCIndex  As Integer

    Dim QuestSlot As Byte
    
    Dim i         As Long
    
    Dim NextQuest As Integer
    
    NPCIndex = UserList(UserIndex).flags.TargetNPC
    
    '¿El NPC es valido?
    If NPCIndex < 1 Then
        userAceptaquest = False
        Exit Function
    End If
    
    'Esta el personaje en la distancia correcta?
    If Distancia(UserList(UserIndex).Pos, Npclist(NPCIndex).Pos) > 5 Then
        Call WriteConsoleMsg(UserIndex, "Estas demasiado lejos.", FontTypeNames.FONTTYPE_INFO)
        userAceptaquest = False
        Exit Function
'
    End If
    
    'Esta el user muerto?
    If UserList(UserIndex).flags.Muerto = 1 Then
        Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
        userAceptaquest = False
        Exit Function

    End If
    
    'Buscamos la siguiente quest que da el NPC (Cadena de quests)
    NextQuest = BuscarSiguienteQuest(UserIndex, NPCIndex)
    
    'Agregamos la quest.
    With UserList(UserIndex).QuestStats.Quests(NextQuest)
    
        'Solicitamos un slot libre para quest aceptadas (asi comprobamos tambien si puede aceptar mas quest)
        QuestSlot = MaxQuestsAceptadas(UserIndex)
        
        'Si obtuvimos -1, es que no hay slots libres y cancelamos (volvemos a comprobar)
        If QuestSlot < 0 Then
            Call WriteConsoleMsg(UserIndex, "Estas haciendo demasiadas misiones. Vuelve cuando hayas completado alguna.", FontTypeNames.FONTTYPE_INFO)
            userAceptaquest = False
            Exit Function
        End If
    
        'Añadimos la quest a la lista de seguimiento
        If AddQuestEnCurso(UserIndex, QuestSlot, NextQuest) = False Then
            Call WriteConsoleMsg(UserIndex, "Estas haciendo demasiadas misiones. Vuelve cuando hayas completado alguna.", FontTypeNames.FONTTYPE_INFO)
            userAceptaquest = False
            Exit Function
        End If
        
        'Establecemos el estado de la quest en curso
        .QuestStatus = eStatusQuest.EnCurso
        
        'Si la quest requiere matar NPC preparamos el Array
        If QuestList(NextQuest).RequiredNPCs Then
            ReDim .NPCsKilled(1 To QuestList(NextQuest).RequiredNPCs)
            
            For i = 1 To QuestList(NextQuest).RequiredNPCs
                .NPCsKilled(i) = 0
            Next i
            
        End If
        
        Call WriteConsoleMsg(UserIndex, "Has aceptado la mision " & Chr(34) & QuestList(NextQuest).Nombre & Chr(34) & ".", FontTypeNames.FONTTYPE_INFO)
        
        'Actualizamos el simbolito del NPC
        Call ActualizarNPCQuest(UserIndex, NPCIndex)
        
    End With
    
    userAceptaquest = True
    
    Exit Function
    
ErrorHandler:
    userAceptaquest = False
    Call LogError("Error userAceptaquest: " & Err.Number & " - " & Err.description)
    
End Function

