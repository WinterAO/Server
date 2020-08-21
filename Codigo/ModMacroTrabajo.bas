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
    Sastreando = 22
    Plantitas = 23
    CreandoPotis = 24
End Enum

Public Function PuedePescar(ByVal UserIndex As Integer) As Boolean
'************************************
'Autor: Lorwik
'Requisitos para pescar
'************************************

    Dim DummyINT As Integer

    With UserList(UserIndex)
    
        DummyINT = .Invent.WeaponEqpObjIndex
                
        If DummyINT = 0 Then
            Call WriteConsoleMsg(UserIndex, "Necesitas una caña o una red para atrapar peces.", FontTypeNames.FONTTYPE_INFOBOLD)
            Call DejardeTrabajar(UserIndex)
            PuedePescar = False
            Exit Function
        End If
        
        If .flags.invisible = 1 Or .flags.Oculto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡Estas Invisible!", FontTypeNames.FONTTYPE_INFOBOLD)
            Call DejardeTrabajar(UserIndex)
            PuedePescar = False
            Exit Function
        End If
        
        If .Stats.MinSta <= 5 Then
            Call WriteConsoleMsg(UserIndex, "Te encuentras demasiado cansado.", FontTypeNames.FONTTYPE_INFOBOLD)
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(SND_PESCAR, .Pos.X, .Pos.Y))
            Call DejardeTrabajar(UserIndex)
            PuedePescar = False
            Exit Function
        End If
                
        If MapData(.Pos.Map, .Pos.X, .Pos.Y).Trigger = 1 Then
            Call WriteConsoleMsg(UserIndex, "No puedes pescar desde donde te encuentras.", FontTypeNames.FONTTYPE_INFOBOLD)
            Call DejardeTrabajar(UserIndex)
            PuedePescar = False
            Exit Function
        End If
        
        PuedePescar = True
        
    End With
    
End Function

Public Function PuedeExtraer(ByVal UserIndex As Integer, ByVal Skill As Byte) As Boolean
'************************************
'Autor: Lorwik
'Requisitos para Extraer recursos de forma pasiva
'************************************

    With UserList(UserIndex)
    
        'Check interval
        If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Function
        
        If ConoceProfesion(UserIndex, Skill) < 0 Then
            Call WriteConsoleMsg(UserIndex, "No conoces esa profesion.", FontTypeNames.FONTTYPE_INFOBOLD)
            PuedeExtraer = False
            Exit Function
        End If
                
        If .Invent.WeaponEqpObjIndex = 0 Then
            Call WriteConsoleMsg(UserIndex, "Deberías equiparte la herramienta.", FontTypeNames.FONTTYPE_INFOBOLD)
            PuedeExtraer = False
            Exit Function
        End If
                
        If ObjData(.Invent.WeaponEqpObjIndex).Herramienta.Profesion <> Skill Then
            Call DejardeTrabajar(UserIndex)
            PuedeExtraer = False
            Exit Function
        End If
        
        
        If MapInfo(UserList(UserIndex).Pos.Map).Pk = False Then
            Call WriteConsoleMsg(UserIndex, "No puedes extraer recursos dentro de la ciudad.", FontTypeNames.FONTTYPE_INFO)
            Call DejardeTrabajar(UserIndex)
            PuedeExtraer = False
            Exit Function
        End If
        
        If .flags.invisible = 1 Or .flags.Oculto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡Estas Invisible!", FontTypeNames.FONTTYPE_INFOBOLD)
            Call DejardeTrabajar(UserIndex)
            PuedeExtraer = False
            Exit Function
        End If
        
        If .Stats.MinSta <= 5 Then
            Call WriteConsoleMsg(UserIndex, "Te encuentras demasiado cansado.", FontTypeNames.FONTTYPE_INFOBOLD)
            Call DejardeTrabajar(UserIndex)
            PuedeExtraer = False
            Exit Function
        End If
    
        PuedeExtraer = True
    End With
    
End Function

Public Function PuedeLingotear(ByVal UserIndex As Integer) As Boolean
'************************************
'Autor: Lorwik
'Requisitos para lingotear
'************************************

    With UserList(UserIndex)
    
        If ConoceProfesion(UserIndex, eSkill.Mineria) < 0 Then
            Call WriteConsoleMsg(UserIndex, "No conoces esa profesion.", FontTypeNames.FONTTYPE_INFOBOLD)
            PuedeLingotear = False
            Exit Function
        End If
    
        If .flags.Equitando Then
            Call WriteConsoleMsg(UserIndex, "No puedes fundir minerales estando montado.", FontTypeNames.FONTTYPE_INFOBOLD)
            Call DejardeTrabajar(UserIndex)
            PuedeLingotear = False
            Exit Function
        End If
        
        If .flags.invisible = 1 Or .flags.Oculto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡Estas Invisible!", FontTypeNames.FONTTYPE_INFOBOLD)
            Call DejardeTrabajar(UserIndex)
            PuedeLingotear = False
            Exit Function
        End If
        
        If .Stats.UserSkills(eSkill.Mineria) < 5 Then
            Call WriteConsoleMsg(UserIndex, "¡No tienes conocimientos en esa profesion!", FontTypeNames.FONTTYPE_INFOBOLD)
            Call DejardeTrabajar(UserIndex)
            PuedeLingotear = False
            Exit Function
        End If
                    
        'Check there is a proper item there
        If .flags.TargetObj > 0 Then
            If ObjData(.flags.TargetObj).OBJType = eOBJType.otFragua Then
            'Validate other items
                If .flags.TargetObjInvSlot < 1 Or .flags.TargetObjInvSlot > MAX_INVENTORY_SLOTS Then
                    Call WriteConsoleMsg(UserIndex, "No tienes mas espacio en tu inventario.", FontTypeNames.FONTTYPE_INFOBOLD)
                    Call DejardeTrabajar(UserIndex)
                    PuedeLingotear = False
                    Exit Function
                End If
                            
                ''chequeamos que no se zarpe duplicando oro
                If .Invent.Object(.flags.TargetObjInvSlot).ObjIndex <> .flags.TargetObjInvIndex Then
                    If .Invent.Object(.flags.TargetObjInvSlot).ObjIndex = 0 Or .Invent.Object(.flags.TargetObjInvSlot).Amount = 0 Then
                        Call DejardeTrabajar(UserIndex)
                        PuedeLingotear = False
                        Exit Function
                    End If
                                
                                ''FUISTE
                    Call WriteErrorMsg(UserIndex, "Has sido expulsado por el sistema anti cheats.")
                    Call FlushBuffer(UserIndex)
                    Call CloseSocket(UserIndex)
                    Exit Function
                End If
                
                'Puede trabajar ;)
                PuedeLingotear = True
                
            Else
                Call WriteConsoleMsg(UserIndex, "No hay ninguna fragua allí.", FontTypeNames.FONTTYPE_INFOBOLD)
                Call DejardeTrabajar(UserIndex)
                PuedeLingotear = False
                Exit Function
            End If
        Else
            Call WriteConsoleMsg(UserIndex, "No hay ninguna fragua allí.", FontTypeNames.FONTTYPE_INFOBOLD)
            Call DejardeTrabajar(UserIndex)
            PuedeLingotear = False
            Exit Function
        End If
                
    End With

End Function

Private Function PuedeCarpinteria(ByVal UserIndex As Integer, ByVal Cantidad As Integer, ByVal Item As Integer) As Boolean
'************************************
'Autor: Lorwik
'Requisitos para construir carpinteria
'************************************
    Dim SlotProfesion As Integer

    With UserList(UserIndex)
    
        If .flags.Trabajando <> eSkill.Carpinteria Then
            PuedeCarpinteria = False
            Exit Function
        End If
    
        SlotProfesion = ConoceProfesion(UserIndex, eSkill.Carpinteria)
        
        If SlotProfesion < 0 Then
            Call WriteConsoleMsg(UserIndex, "No conoces esa profesion.", FontTypeNames.FONTTYPE_INFOBOLD)
            PuedeCarpinteria = False
            Exit Function
        End If
        
        '¿Intento de Hack?
        If Not TieneReceta(Item, UserIndex, SlotProfesion) Then
            PuedeCarpinteria = False
            Exit Function
        End If
        
        '¿El item es inferior a 0 (un item invalido?
        If Item < 1 Then
            Call DejardeTrabajar(UserIndex)
            PuedeCarpinteria = False
            Exit Function
        End If
           
        '¿Ese objeto requiere 0 en skills?
        If ObjData(Item).SkCarpinteria = 0 Then
            Call DejardeTrabajar(UserIndex)
            PuedeCarpinteria = False
            Exit Function
        End If
        
        '¿El contador de objetos pendientes a construir llego a 0?
        If .flags.MacroCountObj < 1 Then
            Call DejardeTrabajar(UserIndex)
            PuedeCarpinteria = False
            Exit Function
        End If
        
        'El usuario esta invisible u oculto?
        If .flags.invisible = 1 Or .flags.Oculto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡Estas Invisible!", FontTypeNames.FONTTYPE_INFOBOLD)
            Call DejardeTrabajar(UserIndex)
            PuedeCarpinteria = False
            Exit Function
        End If
        
        '¿Tiene materiales para construir el proximo item?
        If Not TieneMateriales(UserIndex, Item) Then
            Call DejardeTrabajar(UserIndex)
            PuedeCarpinteria = False
            Exit Function
        End If
        
        '¿Tiene los skills para construir el item?
        If Not UserList(UserIndex).Stats.UserSkills(eSkill.Carpinteria) >= ObjData(Item).SkCarpinteria Then
            Call DejardeTrabajar(UserIndex)
            PuedeCarpinteria = False
            Exit Function
        End If
        
        '¿Tiene el serrucho equipado?
        If Not UserList(UserIndex).Invent.WeaponEqpObjIndex = SERRUCHO_CARPINTERO Then
            Call DejardeTrabajar(UserIndex)
            PuedeCarpinteria = False
            Exit Function
        End If
        
        PuedeCarpinteria = True
    End With
End Function

Private Function PuedeHerreria(ByVal UserIndex As Integer, ByVal Cantidad As Integer, ByVal Item As Integer) As Boolean
'************************************
'Autor: Lorwik
'Requisitos para construir Herreria
'************************************
    Dim SlotProfesion As Integer
    
    With UserList(UserIndex)
    
        If .flags.Trabajando <> eSkill.herreria Then
            PuedeHerreria = False
            Exit Function
        End If
    
        SlotProfesion = ConoceProfesion(UserIndex, eSkill.herreria)
    
        If SlotProfesion < 0 Then
            Call WriteConsoleMsg(UserIndex, "No conoces esa profesion.", FontTypeNames.FONTTYPE_INFOBOLD)
            PuedeHerreria = False
            Exit Function
        End If
        
        '¿Intento de Hack?
        If Not TieneReceta(Item, UserIndex, SlotProfesion) Then
            PuedeHerreria = False
            Exit Function
        End If
        
        '¿El item es inferior a 0 (un item invalido?
        If Item < 1 Then
            Call DejardeTrabajar(UserIndex)
            PuedeHerreria = False
            Exit Function
        End If
        
        '¿Ese objeto requiere 0 en skills?
        If ObjData(Item).SkHerreria = 0 Then
            Call DejardeTrabajar(UserIndex)
            PuedeHerreria = False
            Exit Function
        End If
        
        '¿El contador de objetos pendientes a construir llego a 0?
        If .flags.MacroCountObj < 1 Then
            Call DejardeTrabajar(UserIndex)
            PuedeHerreria = False
            Exit Function
        End If
        
        'El usuario esta invisible u oculto?
        If .flags.invisible = 1 Or .flags.Oculto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡Estas Invisible!", FontTypeNames.FONTTYPE_INFOBOLD)
            Call DejardeTrabajar(UserIndex)
            PuedeHerreria = False
            Exit Function
        End If
        
        If Not PuedeConstruirItemHerrero(UserIndex, Item) Then
            Call WriteConsoleMsg(UserIndex, "No puedes construir ese objeto, te faltan materiales o la habilidad para poder construirlo.", FontTypeNames.FONTTYPE_INFO)
            Call DejardeTrabajar(UserIndex)
            PuedeHerreria = False
            Exit Function
        End If
        
        If Not PuedeConstruirHerreria(Item) Then
            Call DejardeTrabajar(UserIndex)
            PuedeHerreria = False
            Exit Function
        End If
        
        '¿Tiene el martillo equipado?
        If Not UserList(UserIndex).Invent.WeaponEqpObjIndex = MARTILLO_HERRERO Then
            Call DejardeTrabajar(UserIndex)
            PuedeHerreria = False
            Exit Function
        End If
        
        PuedeHerreria = True
    End With
End Function

Private Function PuedeSastreria(ByVal UserIndex As Integer, ByVal Cantidad As Integer, ByVal Item As Integer) As Boolean
'************************************
'Autor: Lorwik
'Requisitos para construir sastreria
'************************************
    Dim SlotProfesion As Integer

    With UserList(UserIndex)
    
        If .flags.Trabajando <> eSkill.Sastreria Then
            PuedeSastreria = False
            Exit Function
        End If
    
        SlotProfesion = ConoceProfesion(UserIndex, eSkill.Sastreria)
        
        If SlotProfesion < 0 Then
            Call WriteConsoleMsg(UserIndex, "No conoces esa profesion.", FontTypeNames.FONTTYPE_INFOBOLD)
            PuedeSastreria = False
            Exit Function
        End If
        
        '¿Intento de Hack?
        If Not TieneReceta(Item, UserIndex, SlotProfesion) Then
            PuedeSastreria = False
            Exit Function
        End If
        
        '¿El item es inferior a 0 (un item invalido?
        If Item < 1 Then
            Call DejardeTrabajar(UserIndex)
            PuedeSastreria = False
            Exit Function
        End If
           
        '¿Ese objeto requiere 0 en skills?
        If ObjData(Item).SkSastreria = 0 Then
            Call DejardeTrabajar(UserIndex)
            PuedeSastreria = False
            Exit Function
        End If
        
        '¿El contador de objetos pendientes a construir llego a 0?
        If .flags.MacroCountObj < 1 Then
            Call DejardeTrabajar(UserIndex)
            PuedeSastreria = False
            Exit Function
        End If
        
        'El usuario esta invisible u oculto?
        If .flags.invisible = 1 Or .flags.Oculto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡Estas Invisible!", FontTypeNames.FONTTYPE_INFOBOLD)
            Call DejardeTrabajar(UserIndex)
            PuedeSastreria = False
            Exit Function
        End If
        
        '¿Tiene materiales para construir el proximo item?
        If Not TieneMateriales(UserIndex, Item) Then
            Call DejardeTrabajar(UserIndex)
            PuedeSastreria = False
            Exit Function
        End If
        
        '¿Tiene los skills para construir el item?
        If Not UserList(UserIndex).Stats.UserSkills(eSkill.Sastreria) >= ObjData(Item).SkSastreria Then
            Call DejardeTrabajar(UserIndex)
            PuedeSastreria = False
            Exit Function
        End If
        
        '¿Tiene el kit de sastreria equipado?
        If Not UserList(UserIndex).Invent.WeaponEqpObjIndex = KIT_DE_COSTURA Then
            Call DejardeTrabajar(UserIndex)
            PuedeSastreria = False
            Exit Function
        End If
        
        PuedeSastreria = True
    End With
End Function

Private Function PuedeAlquimia(ByVal UserIndex As Integer, ByVal Cantidad As Integer, ByVal Item As Integer) As Boolean
'************************************
'Autor: Lorwik
'Requisitos para construir alquimia
'************************************
    Dim SlotProfesion As Integer

    With UserList(UserIndex)
    
        If .flags.Trabajando <> eSkill.Alquimia Then
            PuedeAlquimia = False
            Exit Function
        End If
    
        SlotProfesion = ConoceProfesion(UserIndex, eSkill.Alquimia)
        
        If SlotProfesion < 0 Then
            Call WriteConsoleMsg(UserIndex, "No conoces esa profesion.", FontTypeNames.FONTTYPE_INFOBOLD)
            PuedeAlquimia = False
            Exit Function
        End If
        
        '¿Intento de Hack?
        If Not TieneReceta(Item, UserIndex, SlotProfesion) Then
            PuedeAlquimia = False
            Exit Function
        End If
        
        '¿El item es inferior a 0 (un item invalido?
        If Item < 1 Then
            Call DejardeTrabajar(UserIndex)
            PuedeAlquimia = False
            Exit Function
        End If
           
        '¿Ese objeto requiere 0 en skills?
        If ObjData(Item).SkSastreria = 0 Then
            Call DejardeTrabajar(UserIndex)
            PuedeAlquimia = False
            Exit Function
        End If
        
        '¿El contador de objetos pendientes a construir llego a 0?
        If .flags.MacroCountObj < 1 Then
            Call DejardeTrabajar(UserIndex)
            PuedeAlquimia = False
            Exit Function
        End If
        
        'El usuario esta invisible u oculto?
        If .flags.invisible = 1 Or .flags.Oculto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡Estas Invisible!", FontTypeNames.FONTTYPE_INFOBOLD)
            Call DejardeTrabajar(UserIndex)
            PuedeAlquimia = False
            Exit Function
        End If
        
        '¿Tiene materiales para construir el proximo item?
        If Not TieneMateriales(UserIndex, Item) Then
            Call DejardeTrabajar(UserIndex)
            PuedeAlquimia = False
            Exit Function
        End If
        
        '¿Tiene los skills para construir el item?
        If Not UserList(UserIndex).Stats.UserSkills(eSkill.Alquimia) >= ObjData(Item).SkAlquimia Then
            Call DejardeTrabajar(UserIndex)
            PuedeAlquimia = False
            Exit Function
        End If
        
        '¿Tiene el kit de sastreria equipado?
        If Not UserList(UserIndex).Invent.WeaponEqpObjIndex = OLLA_ALQUIMISTA Then
            Call DejardeTrabajar(UserIndex)
            PuedeAlquimia = False
            Exit Function
        End If
        
        PuedeAlquimia = True
    End With
End Function

Public Sub ComenzarCrafteo(ByVal UserIndex As Integer, ByVal Item As Long, ByVal Cantidad As Integer, ByVal Profesion As Byte)
'************************************************
'Autor: Lorwik
'Ultima modificacion: 21/08/2020
'Gestiona todo para poder empezar a caftear
'************************************************

    With UserList(UserIndex)
    
        If Not IntervaloPermiteTrabajar(UserIndex) Then Exit Sub
    
        Select Case Profesion
        
            Case eSkill.herreria
            
                If ObjData(Item).SkHerreria = 0 Then Exit Sub
            
            Case eSkill.Carpinteria
                If ObjData(Item).SkCarpinteria = 0 Then Exit Sub
                
            Case eSkill.Sastreria
                If ObjData(Item).SkSastreria = 0 Then Exit Sub
            
            Case eSkill.Alquimia
                If ObjData(Item).SkAlquimia = 0 Then Exit Sub
            
        End Select
        
        'Comprobamos que no se encuentra trabajando, para prevenir bugs y hacks
        If .flags.MacroTrabajo = 0 Then
            .flags.MacroTrabajaObj = Item
            .flags.MacroCountObj = Cantidad
            .flags.MacroTrabajo = Profesion
            Call WriteConsoleMsg(UserIndex, "Comienzas a trabajar.", FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(UserIndex, "Ya te encuentras trabajando.", FontTypeNames.FONTTYPE_INFO)
        End If

    End With
    
End Sub

Public Sub DejardeTrabajar(ByVal UserIndex)
'************************************************
'Autor: Lorwik
'Ultima modificacion: 28/03/2020
'Si el usuario esta trabajando, deja de trabajar reseteando los flags del macro
'************************************************

    With UserList(UserIndex)
        'Comprobamos por si acaso que este trabajando
        If .flags.MacroTrabajo > 0 Then
            Call WriteStopWorking(UserIndex)
            .flags.MacroTrabajo = 0
            .flags.MacroTrabajaObj = 0
            .flags.MacroCountObj = 0
            .Counters.Trabajando = 0
        End If
    End With

End Sub

Public Sub MacroTrabajo(ByVal UserIndex As Integer, ByRef Tarea As eMacroTrabajo)
'************************************
'Autor: Lorwik
'Inicia la actividad
'************************************
Debug.Print Tarea
    With UserList(UserIndex)
        Select Case Tarea
        
            'Pesca con caña
            Case eMacroTrabajo.PESCAR
                If PuedePescar(UserIndex) Then
                    Call DoPescar(UserIndex, False)
                Else
                    Call DejardeTrabajar(UserIndex)
                End If
            
            'Pesca con red
            Case eMacroTrabajo.PescarRed
                If PuedePescar(UserIndex) Then
                    Call DoPescar(UserIndex, True)
                Else
                    Call DejardeTrabajar(UserIndex)
                End If
                    
            'Mineria, Talar
            Case eMacroTrabajo.Minando, eMacroTrabajo.Talando, eMacroTrabajo.Plantitas
                If PuedeExtraer(UserIndex, Tarea) Then
                    Call DoExtraer(UserIndex, Tarea)
                Else
                    Call DejardeTrabajar(UserIndex)
                End If
                
            'Lingotear
            Case eMacroTrabajo.Lingotear
                If PuedeLingotear(UserIndex) Then
                    Call FundirMineral(UserIndex)
                Else
                    Call DejardeTrabajar(UserIndex)
                End If

            'Carpinteria
            Case eMacroTrabajo.Carpinteando
                If PuedeCarpinteria(UserIndex, .flags.MacroCountObj, .flags.MacroTrabajaObj) And .flags.MacroCountObj > 0 Then
                    Call CarpinteroConstruirItem(UserIndex, .flags.MacroTrabajaObj)
                    .flags.MacroCountObj = .flags.MacroCountObj - 1 'Restamos en 1 a la cantidad de objetos que queremos construir
                Else
                    Call DejardeTrabajar(UserIndex)
                End If
                
            'Herreria
            Case eMacroTrabajo.Herreando
                If PuedeHerreria(UserIndex, .flags.MacroCountObj, .flags.MacroTrabajaObj) And .flags.MacroCountObj > 0 Then
                    Call HerreroConstruirItem(UserIndex, .flags.MacroTrabajaObj)
                    .flags.MacroCountObj = .flags.MacroCountObj - 1 'Restamos en 1 a la cantidad de objetos que queremos construir
                Else
                    Call DejardeTrabajar(UserIndex)
                End If
                
            'Sastreria
            Case eMacroTrabajo.Sastreando
                If PuedeSastreria(UserIndex, .flags.MacroCountObj, .flags.MacroTrabajaObj) And .flags.MacroCountObj > 0 Then
                    Call SastreConstruirItem(UserIndex, .flags.MacroTrabajaObj)
                    .flags.MacroCountObj = .flags.MacroCountObj - 1 'Restamos en 1 a la cantidad de objetos que queremos construir
                Else
                    Call DejardeTrabajar(UserIndex)
                End If
                
            'Alquimia
            Case eMacroTrabajo.CreandoPotis
                If PuedeAlquimia(UserIndex, .flags.MacroCountObj, .flags.MacroTrabajaObj) And .flags.MacroCountObj > 0 Then
                    Call AlquimistaConstruirItem(UserIndex, .flags.MacroTrabajaObj)
                    .flags.MacroCountObj = .flags.MacroCountObj - 1 'Restamos en 1 a la cantidad de objetos que queremos construir
                Else
                    Call DejardeTrabajar(UserIndex)
                End If
            
        End Select
    End With
        
End Sub
