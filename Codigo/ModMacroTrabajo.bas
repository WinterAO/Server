Attribute VB_Name = "ModMacroTrabajo"
'********************************Modulo Macro**********************************
'Author: Lorwik
'Last Modification: 22/03/2020
'Control asistido de trabajo.
'22/03/2020: Implementado nuevo sistema de macro de trabajo que le da todo el _
control al server.
'******************************************************************************

Option Explicit

Public Enum eMacroTrabajo '(El 0 es no activado)
    Ninguno = 0
    Lingotear = 1
    PescarRed = 2
    'DEBE y Coincide con el numero de los skills:
    Talando = 17
    PESCAR = 18
    Minando = 19
    Carpinteando = 20
    Herreando = 21
End Enum

Public Function PuedePescar(ByVal userIndex As Integer) As Boolean
'************************************
'Autor: Lorwik
'Requisitos para pescar
'************************************

    Dim DummyINT As Integer

    With UserList(userIndex)
    
        DummyINT = .Invent.WeaponEqpObjIndex
                
        If DummyINT = 0 Then
            Call WriteConsoleMsg(userIndex, "Necesitas una caña o una red para atrapar peces.", FontTypeNames.FONTTYPE_INFOBOLD)
            Call DejardeTrabajar(userIndex)
            PuedePescar = False
            Exit Function
        End If
        
        If .flags.invisible = 1 Or .flags.Oculto = 1 Then
            Call WriteConsoleMsg(userIndex, "¡Estas Invisible!", FontTypeNames.FONTTYPE_INFOBOLD)
            Call DejardeTrabajar(userIndex)
            PuedePescar = False
            Exit Function
        End If
        
        If .Stats.MinSta <= 5 Then
            Call WriteConsoleMsg(userIndex, "Te encuentras demasiado cansado.", FontTypeNames.FONTTYPE_INFOBOLD)
            Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessagePlayWave(SND_PESCAR, .Pos.X, .Pos.Y))
            Call DejardeTrabajar(userIndex)
            PuedePescar = False
            Exit Function
        End If
                
        If MapData(.Pos.Map, .Pos.X, .Pos.Y).Trigger = 1 Then
            Call WriteConsoleMsg(userIndex, "No puedes pescar desde donde te encuentras.", FontTypeNames.FONTTYPE_INFOBOLD)
            Call DejardeTrabajar(userIndex)
            PuedePescar = False
            Exit Function
        End If
        
        PuedePescar = True
        
    End With
    
End Function

Public Function PuedeExtraer(ByVal userIndex As Integer, ByVal Skill As Byte) As Boolean
'************************************
'Autor: Lorwik
'Requisitos para Extraer recursos de forma pasiva
'************************************

    With UserList(userIndex)
    
        'Check interval
        If Not IntervaloPermiteTrabajar(userIndex) Then Exit Function
                
        If .Invent.WeaponEqpObjIndex = 0 Then
            Call WriteConsoleMsg(userIndex, "Deberías equiparte la herramienta.", FontTypeNames.FONTTYPE_INFOBOLD)
            PuedeExtraer = False
            Exit Function
        End If
                
        If ObjData(.Invent.WeaponEqpObjIndex).Herramienta.Profesion <> Skill Then
            Call DejardeTrabajar(userIndex)
            PuedeExtraer = False
            Exit Function
        End If
        
        
        If MapInfo(UserList(userIndex).Pos.Map).Pk = False Then
            Call WriteConsoleMsg(userIndex, "No puedes extraer recursos dentro de la ciudad.", FontTypeNames.FONTTYPE_INFO)
            Call DejardeTrabajar(userIndex)
            PuedeExtraer = False
            Exit Function
        End If
        
        If .flags.invisible = 1 Or .flags.Oculto = 1 Then
            Call WriteConsoleMsg(userIndex, "¡Estas Invisible!", FontTypeNames.FONTTYPE_INFOBOLD)
            Call DejardeTrabajar(userIndex)
            PuedeExtraer = False
            Exit Function
        End If
        
        If .Stats.MinSta <= 5 Then
            Call WriteConsoleMsg(userIndex, "Te encuentras demasiado cansado.", FontTypeNames.FONTTYPE_INFOBOLD)
            Call DejardeTrabajar(userIndex)
            PuedeExtraer = False
            Exit Function
        End If
    
        PuedeExtraer = True
    End With
    
End Function

Public Function PuedeLingotear(ByVal userIndex As Integer) As Boolean
'************************************
'Autor: Lorwik
'Requisitos para lingotear
'************************************

    With UserList(userIndex)
    
        If .flags.Equitando Then
            Call WriteConsoleMsg(userIndex, "No puedes fundir minerales estando montado.", FontTypeNames.FONTTYPE_INFOBOLD)
            Call DejardeTrabajar(userIndex)
            PuedeLingotear = False
            Exit Function
        End If
        
        If .flags.invisible = 1 Or .flags.Oculto = 1 Then
            Call WriteConsoleMsg(userIndex, "¡Estas Invisible!", FontTypeNames.FONTTYPE_INFOBOLD)
            Call DejardeTrabajar(userIndex)
            PuedeLingotear = False
            Exit Function
        End If
        
        If .Stats.UserSkills(eSkill.Mineria) < 5 Then
            Call WriteConsoleMsg(userIndex, "¡No tienes conocimientos en esa profesion!", FontTypeNames.FONTTYPE_INFOBOLD)
            Call DejardeTrabajar(userIndex)
            PuedeLingotear = False
            Exit Function
        End If
                    
        'Check there is a proper item there
        If .flags.TargetObj > 0 Then
            If ObjData(.flags.TargetObj).OBJType = eOBJType.otFragua Then
            'Validate other items
                If .flags.TargetObjInvSlot < 1 Or .flags.TargetObjInvSlot > MAX_INVENTORY_SLOTS Then
                    Call WriteConsoleMsg(userIndex, "No tienes mas espacio en tu inventario.", FontTypeNames.FONTTYPE_INFOBOLD)
                    Call DejardeTrabajar(userIndex)
                    PuedeLingotear = False
                    Exit Function
                End If
                            
                ''chequeamos que no se zarpe duplicando oro
                If .Invent.Object(.flags.TargetObjInvSlot).ObjIndex <> .flags.TargetObjInvIndex Then
                    If .Invent.Object(.flags.TargetObjInvSlot).ObjIndex = 0 Or .Invent.Object(.flags.TargetObjInvSlot).Amount = 0 Then
                        Call DejardeTrabajar(userIndex)
                        PuedeLingotear = False
                        Exit Function
                    End If
                                
                                ''FUISTE
                    Call WriteErrorMsg(userIndex, "Has sido expulsado por el sistema anti cheats.")
                    Call FlushBuffer(userIndex)
                    Call CloseSocket(userIndex)
                    Exit Function
                End If
                
                'Puede trabajar ;)
                PuedeLingotear = True
                
            Else
                Call WriteConsoleMsg(userIndex, "No hay ninguna fragua allí.", FontTypeNames.FONTTYPE_INFOBOLD)
                Call DejardeTrabajar(userIndex)
                PuedeLingotear = False
                Exit Function
            End If
        Else
            Call WriteConsoleMsg(userIndex, "No hay ninguna fragua allí.", FontTypeNames.FONTTYPE_INFOBOLD)
            Call DejardeTrabajar(userIndex)
            PuedeLingotear = False
            Exit Function
        End If
                
    End With

End Function

Private Function PuedeCarpinteria(ByVal userIndex As Integer, ByVal Cantidad As Integer, ByVal Item As Integer) As Boolean
'************************************
'Autor: Lorwik
'Requisitos para construir carpinteria
'************************************

    With UserList(userIndex)
        
        '¿El item es inferior a 0 (un item invalido?
        If Item < 1 Then
            Call DejardeTrabajar(userIndex)
            PuedeCarpinteria = False
            Exit Function
        End If
           
        '¿Ese objeto requiere 0 en skills?
        If ObjData(Item).SkCarpinteria = 0 Then
            Call DejardeTrabajar(userIndex)
            PuedeCarpinteria = False
            Exit Function
        End If
        
        '¿El contador de objetos pendientes a construir llego a 0?
        If .flags.MacroCountObj < 1 Then
            Call DejardeTrabajar(userIndex)
            PuedeCarpinteria = False
            Exit Function
        End If
        
        'El usuario esta invisible u oculto?
        If .flags.invisible = 1 Or .flags.Oculto = 1 Then
            Call WriteConsoleMsg(userIndex, "¡Estas Invisible!", FontTypeNames.FONTTYPE_INFOBOLD)
            Call DejardeTrabajar(userIndex)
            PuedeCarpinteria = False
            Exit Function
        End If
        
        '¿Tiene materiales para construir el proximo item?
        If Not CarpinteroTieneMateriales(userIndex, Item) Then
            Call DejardeTrabajar(userIndex)
            PuedeCarpinteria = False
            Exit Function
        End If
        
        '¿Tiene los skills para construir el item?
        If Not UserList(userIndex).Stats.UserSkills(eSkill.Carpinteria) >= ObjData(Item).SkCarpinteria Then
            Call DejardeTrabajar(userIndex)
            PuedeCarpinteria = False
            Exit Function
        End If
        
        '¿Puede construir el item?
        If Not PuedeConstruirCarpintero(Item) Then
            Call DejardeTrabajar(userIndex)
            PuedeCarpinteria = False
            Exit Function
        End If
        
        '¿Tiene el serrucho equipado?
        If Not UserList(userIndex).Invent.WeaponEqpObjIndex = SERRUCHO_CARPINTERO Then
            Call DejardeTrabajar(userIndex)
            PuedeCarpinteria = False
            Exit Function
        End If
        
        PuedeCarpinteria = True
    End With
End Function

Private Function PuedeHerreria(ByVal userIndex As Integer, ByVal Cantidad As Integer, ByVal Item As Integer) As Boolean
'************************************
'Autor: Lorwik
'Requisitos para construir Herreria
'************************************

    With UserList(userIndex)
    
        '¿El item es inferior a 0 (un item invalido?
        If Item < 1 Then
            Call DejardeTrabajar(userIndex)
            PuedeHerreria = False
            Exit Function
        End If
        
        '¿Ese objeto requiere 0 en skills?
        If ObjData(Item).SkHerreria = 0 Then
            Call DejardeTrabajar(userIndex)
            PuedeHerreria = False
            Exit Function
        End If
        
        '¿El contador de objetos pendientes a construir llego a 0?
        If .flags.MacroCountObj < 1 Then
            Call DejardeTrabajar(userIndex)
            PuedeHerreria = False
            Exit Function
        End If
        
        'El usuario esta invisible u oculto?
        If .flags.invisible = 1 Or .flags.Oculto = 1 Then
            Call WriteConsoleMsg(userIndex, "¡Estas Invisible!", FontTypeNames.FONTTYPE_INFOBOLD)
            Call DejardeTrabajar(userIndex)
            PuedeHerreria = False
            Exit Function
        End If
        
        If Not PuedeConstruirItemHerrero(userIndex, Item) Then
            Call DejardeTrabajar(userIndex)
            PuedeHerreria = False
            Exit Function
        End If
        
        If Not PuedeConstruirHerreria(Item) Then
            Call DejardeTrabajar(userIndex)
            PuedeHerreria = False
            Exit Function
        End If
        
        '¿Tiene el martillo equipado?
        If Not UserList(userIndex).Invent.WeaponEqpObjIndex = MARTILLO_HERRERO Then
            Call DejardeTrabajar(userIndex)
            PuedeHerreria = False
            Exit Function
        End If
        
        PuedeHerreria = True
    End With
End Function

Public Sub DejardeTrabajar(ByVal userIndex)
'************************************************
'Autor: Lorwik
'Ultima modificacion: 28/03/2020
'Si el usuario esta trabajando, deja de trabajar reseteando los flags del macro
'************************************************

    With UserList(userIndex)
        'Comprobamos por si acaso que este trabajando
        If .flags.MacroTrabajo > 0 Then
            Call WriteStopWorking(userIndex)
            .flags.MacroTrabajo = 0
            .flags.MacroTrabajaObj = 0
            .flags.MacroCountObj = 0
            .Counters.Trabajando = 0
        End If
    End With

End Sub

Public Sub MacroTrabajo(ByVal userIndex As Integer, ByRef Tarea As eMacroTrabajo)
'************************************
'Autor: Lorwik
'Inicia la actividad
'************************************
Debug.Print Tarea
    With UserList(userIndex)
        Select Case Tarea
        
            'Pesca con caña
            Case eMacroTrabajo.PESCAR
                If PuedePescar(userIndex) Then _
                    Call DoPescar(userIndex, False)
            
            'Pesca con red
            Case eMacroTrabajo.PescarRed
                If PuedePescar(userIndex) Then _
                    Call DoPescar(userIndex, True)
                    
            'Mineria, Talar
            Case eMacroTrabajo.Minando, eMacroTrabajo.Talando
                If PuedeExtraer(userIndex, Tarea) Then _
                    Call DoExtraer(userIndex, Tarea)
                    
            'Lingotear
            Case eMacroTrabajo.Lingotear
                If PuedeLingotear(userIndex) Then _
                    Call FundirMineral(userIndex)

            'Carpinteria
            Case eMacroTrabajo.Carpinteando
                If PuedeCarpinteria(userIndex, .flags.MacroCountObj, .flags.MacroTrabajaObj) And .flags.MacroCountObj > 0 Then
                    Call CarpinteroConstruirItem(userIndex, .flags.MacroTrabajaObj)
                    .flags.MacroCountObj = .flags.MacroCountObj - 1 'Restamos en 1 a la cantidad de objetos que queremos construir
                End If
                
            'Herreria
            Case eMacroTrabajo.Herreando
                If PuedeHerreria(userIndex, .flags.MacroCountObj, .flags.MacroTrabajaObj) And .flags.MacroCountObj > 0 Then
                    Call HerreroConstruirItem(userIndex, .flags.MacroTrabajaObj)
                    .flags.MacroCountObj = .flags.MacroCountObj - 1 'Restamos en 1 a la cantidad de objetos que queremos construir
                End If
            
        End Select
    End With
        
End Sub


