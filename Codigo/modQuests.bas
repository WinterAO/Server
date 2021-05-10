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
Public NumQuests As Integer                    'Num de quest dateadas actualmente
 
Public Sub accionUseraNPCQuest(ByVal UserIndex As Integer, ByVal NPCIndex As Integer)
'**************************************
'Autor: Lorwik
'Fecha 09/05/2021
'Descripcion: El usuario interactuo con un NPC que da quest
'**************************************
    On Error GoTo ErrorHandler
    
    Dim QuestIndex As Integer
    Dim Questslot As Integer
    Dim NPCQuestslot As Integer
    
    With Npclist(NPCIndex)
    
        '¿El NPC tiene quest para dar?
        If Npclist(NPCIndex).NumQuest = 0 Then
            Call WriteChatOverHead(UserIndex, "No tengo ningun encargo para ti.", .Char.CharIndex, vbWhite)
            Exit Sub
    
        End If
    
        NPCQuestslot = buscarSiguienteQuest(UserIndex, NPCIndex, Questslot)
        
        '¿Llego un numero superior al numero de quest max. que puede dar el NPC?
        If NPCQuestslot > .NumQuest Then
        
            'Establecemos el slot a la ultima quest que puede dar el NPC
            NPCQuestslot = .NumQuest
            QuestIndex = .QuestNumber(NPCQuestslot)
            
            '¿Es una mision repetible?
            If QuestList(QuestIndex).Repetible Then

                '¿Llego el dia en el que la puede repetir?
                If UserList(UserIndex).QuestStats.Quests(Questslot).fechaFin + QuestList(QuestIndex).Tiempo > Now() Then
                    Call WriteChatOverHead(UserIndex, "En estos momentos no tengo ningun encargo para ti. Vuelve dentro de " & DateDiff("d", Now(), UserList(UserIndex).QuestStats.Quests(Questslot).fechaFin + QuestList(QuestIndex).Tiempo) & " dias.", .Char.CharIndex, vbWhite)
                    Exit Sub
                    
                End If
                
            Else
                Call WriteChatOverHead(UserIndex, "Ya no tengo ningun encargo para ti.", .Char.CharIndex, vbWhite)
                Exit Sub
                
            End If
        
        End If
        
        QuestIndex = .QuestNumber(NPCQuestslot)
        
        'Comprobamos si la quest esta en la lista de quest en curso
        Questslot = buscarQuestenCurso(UserIndex, QuestIndex)
        
        'Si el usuario ya tiene la quest es por que la quiere entregar.
        If Questslot > 0 Then
            Call userFinalizaQuest(UserIndex, Questslot)
            Exit Sub
        End If
        
        'REQUISITOS PARA NUEVA MISION
        
        '¿Cumple el requisito en nivel para poder hacerla?
        If UserList(UserIndex).Stats.ELV < QuestList(QuestIndex).RequiredLevel Then
            Call WriteChatOverHead(UserIndex, "Debes ser por lo menos nivel " & QuestList(QuestIndex).RequiredLevel & " para emprender esta mision.", .Char.CharIndex, vbWhite)
            Exit Sub
    
        End If
        
        If QuestList(QuestIndex).RequiredQuest Then
            
            If Not userYaHizoQuest(UserIndex, QuestList(QuestIndex).RequiredQuest) Then
                Call WriteChatOverHead(UserIndex, "Antes debes haber completado la mision " & Chr(34) & QuestList(QuestList(QuestIndex).RequiredQuest).Nombre & Chr(34), .Char.CharIndex, vbWhite)
                Exit Sub
            End If
        
        End If
        
        'Si llego hasta aqui es por que no la tiene, le enviamos los detalles
        UserList(UserIndex).flags.TargetQuest = QuestIndex
        Call WriteQuestDetails(UserIndex, QuestIndex)
    
    End With
    
    Exit Sub
    
ErrorHandler:
    Call LogError("Error accionUseraNPCQuest: " & Err.Number & " - " & Err.description)
    
End Sub
 
Public Sub userAceptaquest(ByVal UserIndex As Integer)
'**************************************
'Autor: Lorwik
'Fecha 09/05/2021
'Descripcion: El usuario Acepta una quest.
'**************************************
    On Error GoTo ErrorHandler

    Dim NPCIndex As Integer
    Dim QuestIndex As Integer
    Dim reQuestSlot As Integer
    Dim Questslot As Integer
    Dim Reintenta As Boolean
    
    With UserList(UserIndex).QuestStats

        NPCIndex = UserList(UserIndex).flags.TargetNPC
        QuestIndex = UserList(UserIndex).flags.TargetQuest
        
        'Breve medida antihack
        'Si recibimos un false es por que no hubo exito en la comparacion y se trata de un posible hack
        If CompararNPCyQuest(UserIndex, NPCIndex, QuestIndex) = False Then Exit Sub
        
        If .nQuestCurso >= MAXUSERQUESTS Then
            Call WriteConsoleMsg(UserIndex, "Estas haciendo demasiadas misiones. Vuelve cuando hayas completado o abandonado alguna.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        'Buscamos en la lista de quest completadas para saber si la esta repitiendo
        reQuestSlot = userYaHizoQuest(UserIndex, QuestIndex)
        If reQuestSlot > 0 Then
            Questslot = .QuestDone(reQuestSlot)
            Call reOrdenarCompletados(UserIndex, reQuestSlot)
            Reintenta = True
            
        Else
            'Buscamos en la lista de quest abandonadas para saber si esta retomando una quest
            reQuestSlot = buscarQuestAbandonada(UserIndex, QuestIndex)
            If reQuestSlot > 0 Then
                Questslot = .QuestLeave(reQuestSlot)
                Call reOrdenarAbandonados(UserIndex, reQuestSlot)
                Reintenta = True
                
            End If
            
        End If
        
        'Añadimos la quest a seguimiento
        .nQuestCurso = .nQuestCurso + 1
        ReDim Preserve .QuestEnCurso(1 To .nQuestCurso) As Integer
        
        'Si es la primera vez que hace la quest
        If Reintenta = False Then
            .TotalQuest = .TotalQuest + 1
            ReDim Preserve .Quests(1 To .TotalQuest) As tUserQuest
            Questslot = .TotalQuest
        End If

        .QuestEnCurso(.nQuestCurso) = Questslot
        .Quests(Questslot).QuestIndex = QuestIndex
        .Quests(Questslot).QuestStatus = eStatusQuest.EnCurso
        .Quests(Questslot).fechaFin = Now()
        
        'Si la quest requiere matar NPC preparamos el Array
        If QuestList(QuestIndex).RequiredNPCs Then ReDim .Quests(Questslot).NPCsKilled(1 To QuestList(QuestIndex).RequiredNPCs)
        If QuestList(QuestIndex).RequiredTargetNPCs Then ReDim .Quests(Questslot).NPCsTarget(1 To QuestList(QuestIndex).RequiredTargetNPCs)
        
        Call WriteConsoleMsg(UserIndex, "Has aceptado la mision " & Chr(34) & QuestList(QuestIndex).Nombre & Chr(34) & ".", FontTypeNames.FONTTYPE_INFO)
        
        'Actualizamos el simbolito del NPC
        'Call ActualizarNPCQuest(UserIndex, NPCIndex)

    End With
    
    Exit Sub
    
ErrorHandler:
    Call LogError("Error userAceptaquest: " & Err.Number & " - " & Err.description)
    
End Sub
 
Public Sub userAbandonaQuest(ByVal UserIndex As Integer, ByVal QuestCursoSlot As Integer)
'**************************************
'Autor: Lorwik
'Fecha 09/05/2021
'Descripcion:
'**************************************
    On Error GoTo ErrorHandler
    
    Dim Questslot As Integer
    
    With UserList(UserIndex).QuestStats

        Questslot = .QuestEnCurso(QuestCursoSlot)
        
        'Eliminamos la quest en curso
        Call reOrdenarEnCurso(UserIndex, QuestCursoSlot)
        
        'Agregamos al a lista de abandonados
        .nQuestLeave = .nQuestLeave + 1
        ReDim Preserve .QuestLeave(1 To .nQuestLeave) As Integer
        .QuestLeave(.nQuestLeave) = Questslot
        
        'Cambiamos el estado a no aceptada
        .Quests(Questslot).QuestStatus = eStatusQuest.NoAceptada
        
        Call WriteConsoleMsg(UserIndex, "Has abandonado la mision " & Chr(34) & QuestList(.Quests(Questslot).QuestIndex).Nombre & Chr(34) & ".", FontTypeNames.FONTTYPE_INFO)
        
        'Enviamos la lista de quests actualizada.
        Call WriteQuestListSend(UserIndex)
    
    End With
    
    Exit Sub
    
ErrorHandler:
    Call LogError("Error userAbandonaQuest: " & Err.Number & " - " & Err.description)
    
End Sub

Public Sub userFinalizaQuest(ByVal UserIndex As Integer, ByVal Questslot As Byte)

    '**************************************
    'Autor: Lorwik
    'Fecha 09/05/2021
    'Descripcion:
    '**************************************
    On Error GoTo ErrorHandler
    
    Dim InvSlotsLibres  As Byte
    
    Dim NPCIndex        As Integer

    Dim i               As Byte
    
    With QuestList(UserList(UserIndex).QuestStats.Quests(Questslot).QuestIndex)
    
        NPCIndex = UserList(UserIndex).flags.TargetNPC
  
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

                If .RequiredNPC(i).Amount > UserList(UserIndex).QuestStats.Quests(Questslot).NPCsKilled(i) Then
                    Call WriteChatOverHead(UserIndex, "No has matado todas las criaturas que te he pedido.", Npclist(NPCIndex).Char.CharIndex, vbWhite)
                    Exit Sub

                End If

            Next i

        End If
  
        'Comprobamos que haya hablado con todos los NPC
        If .RequiredTargetNPCs > 0 Then

            For i = 1 To .RequiredTargetNPCs
    
                If .RequiredTargetNPC(i).Amount > UserList(UserIndex).QuestStats.Quests(Questslot).NPCsTarget(i) Then
                    Call WriteChatOverHead(UserIndex, "No has visitado a las personas que te pedi.", Npclist(NPCIndex).Char.CharIndex, vbYellow)
                    Exit Sub
    
                End If
    
            Next i

        End If
  
        'Comprobamos que el usuario tenga espacio para recibir los items.
        If .RewardOBJs > 0 Then

            'Buscamos la cantidad de slots de inventario libres.
            For i = 1 To UserList(UserIndex).CurrentInventorySlots

                If UserList(UserIndex).Invent.Object(i).ObjIndex = 0 Then InvSlotsLibres = InvSlotsLibres + 1
            Next i
      
            'Nos fijamos si entra
            If InvSlotsLibres < .RewardOBJs Then
                Call WriteChatOverHead(UserIndex, "No tienes suficiente espacio en el inventario para recibir la recompensa. Vuelve cuando tengas espacio.", Npclist(NPCIndex).Char.CharIndex, vbYellow)
                Exit Sub

            End If

        End If
        
        'A esta altura ya cumplio los objetivos, entonces se le entregan las recompensas.
        Call WriteConsoleMsg(UserIndex, "Has completado la mision " & Chr(34) & .Nombre & Chr(34) & "!", FontTypeNames.FONTTYPE_INFO)
  
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
                    Call WriteConsoleMsg(UserIndex, "Has recibido " & .RewardOBJ(i).Amount & " " & ObjData(.RewardOBJ(i).ObjIndex).Name & " como recompensa.", FontTypeNames.FONTTYPE_INFO)

                End If

            Next i

        End If
    
    End With
    
    With UserList(UserIndex).QuestStats
    
        'Se agrega que el usuario ya hizo esta quest.
        .Quests(Questslot).QuestStatus = eStatusQuest.Terminada
        .Quests(Questslot).fechaFin = Now()
        .nQuestDone = .nQuestDone + 1
        ReDim Preserve .QuestDone(1 To .nQuestDone) As Integer
        .QuestDone(.nQuestDone) = Questslot
  
        'Eliminamos la quest de la lista de seguimientos
        Call reOrdenarEnCurso(UserIndex, Questslot)

        'Actualizamos el personaje
        Call CheckUserLevel(UserIndex)
        Call UpdateUserInv(True, UserIndex, 0)
        Call WriteUpdateGold(UserIndex)
  
        'Call ActualizarNPCQuest(UserIndex, NPCIndex)
    
    End With
    
    Exit Sub
    
ErrorHandler:
    Call LogError("Error userFinalizaQuest: " & Err.Number & " - " & Err.description)

End Sub
 
Public Function buscarSiguienteQuest(ByVal UserIndex As Integer, ByVal NPCIndex As Integer, ByRef Questslot As Integer) As Integer
'**************************************
'Autor: Lorwik
'Fecha 09/05/2021
'Descripcion: busca una quest no completada por el usuario en la lista del NPC
'**************************************
    On Error GoTo ErrorHandler
    
    Dim Count As Integer
    Dim i As Byte
    Dim j As Integer
    
    With UserList(UserIndex).QuestStats
    
        Count = 1
    
        '¿El usuario tiene quests completadas?
        If .nQuestDone <> 0 Then

            'Recorremos la lista de quest completas del usuario
            For i = 1 To .nQuestDone
                For j = 1 To Npclist(NPCIndex).NumQuest
                    
                    If .Quests(.QuestDone(i)).QuestIndex = Npclist(NPCIndex).QuestNumber(j) Then
                        Count = Count + 1
                        Questslot = .QuestDone(i)
                    End If
                Next j
            Next i
            
        End If
        
        'Devolvemos el slot donde esta la quest en el NPC
         buscarSiguienteQuest = Count
    
    End With
    
    Exit Function
    
ErrorHandler:
    Call LogError("Error userFinalizaQuest: " & Err.Number & " - " & Err.description)
    
End Function

Public Function buscarQuestenCurso(ByVal UserIndex As Integer, ByVal QuestIndex As Integer) As Integer
'**************************************
'Autor: Lorwik
'Fecha 09/05/2021
'Descripcion: busca una quest no completada por el usuario en la lista del NPC
'**************************************
    On Error GoTo ErrorHandler
    
    Dim i As Integer
    
    With UserList(UserIndex).QuestStats
    
        'Si no tiene quest o no esta cursando ninguna...
        If .TotalQuest = 0 Or .nQuestCurso = 0 Then
            buscarQuestenCurso = 0
            Exit Function
        End If
    
        For i = 1 To .nQuestCurso
            'Si encontramos la quest en el listado de quest en curso devolvemos el slot en el que se encuentra.
            If .Quests(.QuestEnCurso(i)).QuestIndex = QuestIndex Then
                buscarQuestenCurso = .QuestEnCurso(i)
                Exit Function
            End If
        Next i
    
    End With
    
    Exit Function
    
ErrorHandler:
    Call LogError("buscarQuestenCurso: " & Err.Number & " - " & Err.description)
    Resume Next
    
End Function

Public Function buscarQuestAbandonada(ByVal UserIndex As Integer, ByVal QuestIndex As Integer) As Integer
'**************************************
'Autor: Lorwik
'Fecha 09/05/2021
'Descripcion: busca una quest abandonada por el usuario en la lista del NPC
'**************************************
    On Error GoTo ErrorHandler
    
    Dim i As Integer
    
    With UserList(UserIndex).QuestStats
    
        'Si no tiene quest o no abandono ninguna...
        If .TotalQuest = 0 Or .nQuestLeave = 0 Then
            buscarQuestAbandonada = 0
            Exit Function
        End If
    
        For i = 1 To .nQuestLeave
            'Si encontramos la quest en el listado de quest abandonadas devolvemos el slot en el que se encuentra.
            If .Quests(.QuestLeave(i)).QuestIndex = QuestIndex Then
                buscarQuestAbandonada = i
                Exit Function
            End If
        Next i
    
    End With
    
    Exit Function
    
ErrorHandler:
    Call LogError("buscarQuestAbandonada: " & Err.Number & " - " & Err.description)
    Resume Next
    
End Function
 
Public Sub ResetQuestStats(ByVal UserIndex As Integer)
    '****************************************************
    'Autor: Lorwik
    'Fecha: 28/06/2020
    'Limpia todos los QuestStats de un usuario
    '****************************************************
    
    Dim i As Integer
 
    With UserList(UserIndex).QuestStats
        
        .nQuestCurso = 0
        .nQuestLeave = 0
        .nQuestDone = 0
        .TotalQuest = 0
        
        Erase .QuestDone
        Erase .QuestEnCurso
        Erase .QuestLeave
        Erase .Quests
        
    End With
End Sub

Private Sub reOrdenarCompletados(ByVal UserIndex As Integer, ByVal Questslot As Integer)

    '****************************************************
    'Autor: Lorwik
    'Fecha: 09/05/2021
    'Elimina una quest de la lista de quest completadas y la reordena
    '****************************************************
    On Error GoTo reOrdenarCompletados_Err

    Dim i As Integer
 
    With UserList(UserIndex).QuestStats

        '¿Es el ultimo elemento de los completados?
        If Questslot = .nQuestDone Then
            'Solo hay 1 elemento en la lista y es el que vamos a eliminar?
            If .nQuestDone = 1 Then
                Erase .QuestDone 'borramos el array
                .nQuestDone = .nQuestDone - 1
                
            Else
                .nQuestDone = .nQuestDone - 1
                ReDim Preserve .QuestDone(1 To .nQuestDone) 'redimensionamos
                
            End If
            
            Exit Sub
        End If

        'Recorremos y reodernamos
        For i = Questslot To .nQuestDone - 1

            .QuestDone(i) = .QuestDone(i + 1)

        Next i
        
        .nQuestDone = .nQuestDone - 1
        ReDim Preserve .QuestDone(1 To .nQuestDone)
    End With
  
    Exit Sub

reOrdenarCompletados_Err:
     Call LogError("reOrdenarCompletados: " & Err.Number & " - " & Err.description)
    Resume Next
  
End Sub

Private Sub reOrdenarAbandonados(ByVal UserIndex As Integer, ByVal Questslot As Integer)

    '****************************************************
    'Autor: Lorwik
    'Fecha: 09/05/2021
    'Elimina una quest de la lista de quest abandonadas y la reordena
    '****************************************************
    On Error GoTo reOrdenarAbandonados_Err

    Dim i As Integer
 
    With UserList(UserIndex).QuestStats

        '¿Es el ultimo elemento de los completados?
        If Questslot = .nQuestLeave Then
            'Solo hay 1 elemento en la lista y es el que vamos a eliminar?
            If .nQuestLeave = 1 Then
                Erase .QuestLeave 'borramos el array
                .nQuestLeave = .nQuestLeave - 1
                
            Else
                .nQuestLeave = .nQuestLeave - 1
                ReDim Preserve .QuestLeave(1 To .nQuestLeave) 'redimensionamos
                
            End If
            
            Exit Sub
        End If

        'Recorremos y reodernamos
        For i = Questslot To .nQuestLeave - 1

            .QuestLeave(i) = .QuestLeave(i + 1)

        Next i
        
        .nQuestLeave = .nQuestLeave - 1
        ReDim Preserve .QuestLeave(1 To .nQuestLeave)
    End With
  
    Exit Sub

reOrdenarAbandonados_Err:
     Call LogError("reOrdenarAbandonados: " & Err.Number & " - " & Err.description)
    Resume Next
  
End Sub

Private Sub reOrdenarEnCurso(ByVal UserIndex As Integer, ByVal Questslot As Integer)

    '****************************************************
    'Autor: Lorwik
    'Fecha: 09/05/2021
    'Elimina una quest de la lista de quest en curso y la reordena
    '****************************************************
    On Error GoTo reOrdenarEnCurso_Err

    Dim i As Integer
 
    With UserList(UserIndex).QuestStats

        '¿Es el ultimo elemento de los completados?
        If Questslot = .nQuestCurso Then
            'Solo hay 1 elemento en la lista y es el que vamos a eliminar?
            If .nQuestCurso = 1 Then
                Erase .QuestEnCurso 'borramos el array
                .nQuestCurso = .nQuestCurso - 1
                
            Else
                .nQuestCurso = .nQuestCurso - 1
                ReDim Preserve .QuestEnCurso(1 To .nQuestCurso) 'redimensionamos
                
            End If
            
            Exit Sub
        End If

        'Recorremos y reodernamos
        For i = Questslot To .nQuestCurso - 1

            .QuestEnCurso(i) = .QuestEnCurso(i + 1)

        Next i
        
        .nQuestCurso = .nQuestCurso - 1
        ReDim Preserve .QuestEnCurso(1 To .nQuestCurso)
    End With
  
    Exit Sub

reOrdenarEnCurso_Err:
     Call LogError("reOrdenarEnCurso: " & Err.Number & " - " & Err.description)
    Resume Next
  
End Sub

Private Function CompararNPCyQuest(ByVal UserIndex As Integer, ByVal NPCIndex As Integer, ByVal QuestIndex As Integer) As Boolean
    '****************************************************
    'Autor: Lorwik
    'Fecha: 09/05/2021
    'Comprueba si el NPC recibido da la quest recibida
    '****************************************************
    On Error GoTo ErrorHandler
    Dim i As Byte
    
    If Npclist(NPCIndex).NumQuest = 0 Then
        CompararNPCyQuest = False
        Exit Function
    End If
    
    For i = 1 To Npclist(NPCIndex).NumQuest
    
        If Npclist(NPCIndex).QuestNumber(i) = QuestIndex Then
            CompararNPCyQuest = True
            Exit Function
        End If
    
    Next i
    
    Debug.Print "No se pudo comparar NPC y Quest."
    CompararNPCyQuest = False
    Exit Function
    
ErrorHandler:
    CompararNPCyQuest = False
    Call LogError("CompararNPCyQuest: " & Err.Number & " - " & Err.description)
End Function

Private Function userYaHizoQuest(ByVal UserIndex As Integer, ByVal QuestIndex As Integer) As Integer
    '****************************************************
    'Autor: Lorwik
    'Fecha: 09/05/2021
    'Comprueba si el usuario ya hizo la quest
    '****************************************************
    On Error GoTo ErrorHandler
    Dim i As Integer
    
    With UserList(UserIndex).QuestStats
    
        If .nQuestDone = 0 Then
            userYaHizoQuest = 0
            Exit Function
        End If
        
        For i = 1 To .nQuestDone
        
            If .Quests(.QuestDone(i)).QuestIndex = QuestIndex Then
                userYaHizoQuest = i 'devolvemos el slot dnetro del array QuestDone
                Exit Function
            End If
            
        Next i
        
        'Si llego aca es por que nunca la hizo
        userYaHizoQuest = 0
    
    End With
    
    Exit Function
    
ErrorHandler:
    userYaHizoQuest = 0
    Call LogError("CuserYaHizoQuest: " & Err.Number & " - " & Err.description)
End Function
