Attribute VB_Name = "Trabajo"
'Argentum Online 0.12.2
'Copyright (C) 2002 Marquez Pablo Ignacio
'
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
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 numero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Codigo Postal 1900
'Pablo Ignacio Marquez

Option Explicit

Private Const GASTO_ENERGIA As Byte = 6

Public Sub DoPermanecerOculto(ByVal userIndex As Integer)

    '********************************************************
    'Autor: Nacho (Integer)
    'Last Modif: 11/19/2009
    'Chequea si ya debe mostrarse
    'Pablo (ToxicWaste): Cambie los ordenes de prioridades porque sino no andaba.
    '13/01/2010: ZaMa - Now hidden on boat pirats recover the proper boat body.
    '13/01/2010: ZaMa - Arreglo condicional para que el bandido camine oculto.
    '********************************************************
    On Error GoTo errHandler

    With UserList(userIndex)
        .Counters.TiempoOculto = .Counters.TiempoOculto - 1

        If .Counters.TiempoOculto <= 0 Then
            If .clase = eClass.Hunter And .Stats.UserSkills(eSkill.Ocultarse) > 90 Then
                If .Invent.ArmourEqpObjIndex = 648 Or .Invent.ArmourEqpObjIndex = 360 Then
                    .Counters.TiempoOculto = IntervaloOculto
                    Exit Sub

                End If

            End If

            .Counters.TiempoOculto = 0
            .flags.Oculto = 0
            
            If .flags.Navegando = 1 Then
                If .clase = eClass.Pirat Then
                    ' Pierde la apariencia de fragata fantasmal
                    Call ToggleBoatBody(userIndex)
                    Call WriteConsoleMsg(userIndex, "Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
                    Call ChangeUserChar(userIndex, .Char.body, .Char.Head, .Char.Heading, NingunArma, NingunEscudo, NingunCasco, NingunAura, NingunAura)

                End If

            Else

                If .flags.invisible = 0 Then
                    Call WriteConsoleMsg(userIndex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
                    Call SetInvisible(userIndex, .Char.CharIndex, False)

                End If

            End If

        End If

    End With
    
    Exit Sub

errHandler:
    Call LogError("Error en Sub DoPermanecerOculto")

End Sub

Public Sub DoOcultarse(ByVal userIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: 13/01/2010 (ZaMa)
    'Pablo (ToxicWaste): No olvidar agregar IntervaloOculto=500 al Server.ini.
    'Modifique la formula y ahora anda bien.
    '13/01/2010: ZaMa - El pirata se transforma en galeon fantasmal cuando se oculta en agua.
    '***************************************************

    On Error GoTo errHandler

    Dim Suerte As Double

    Dim res    As Integer

    Dim Skill  As Integer
    
    With UserList(userIndex)
        Skill = .Stats.UserSkills(eSkill.Ocultarse)
        
        Suerte = (((0.000002 * Skill - 0.0002) * Skill + 0.0064) * Skill + 0.1124) * 100
        
        res = RandomNumber(1, 100)
        
        If res <= Suerte Then
        
            .flags.Oculto = 1
            Suerte = (-0.000001 * (100 - Skill) ^ 3)
            Suerte = Suerte + (0.00009229 * (100 - Skill) ^ 2)
            Suerte = Suerte + (-0.0088 * (100 - Skill))
            Suerte = Suerte + (0.9571)
            Suerte = Suerte * IntervaloOculto
            
            If .clase = eClass.Bandit Then
                .Counters.TiempoOculto = Int(Suerte / 2)
            Else
                .Counters.TiempoOculto = Suerte

            End If
            
            ' No es pirata o es uno sin barca
            If .flags.Navegando = 0 Then
                Call SetInvisible(userIndex, .Char.CharIndex, True)
        
                Call WriteConsoleMsg(userIndex, "Te has escondido entre las sombras!", FontTypeNames.FONTTYPE_INFO)
                ' Es un pirata navegando
            Else
                ' Le cambiamos el body a galeon fantasmal
                .Char.body = iFragataFantasmal
                ' Actualizamos clientes
                Call ChangeUserChar(userIndex, .Char.body, .Char.Head, .Char.Heading, NingunArma, NingunEscudo, NingunCasco, NingunAura, NingunAura)

            End If
            
            Call SubirSkill(userIndex, eSkill.Ocultarse, True)
        Else

            '[CDT 17-02-2004]
            If Not .flags.UltimoMensaje = 4 Then
                Call WriteConsoleMsg(userIndex, "No has logrado esconderte!", FontTypeNames.FONTTYPE_INFO)
                .flags.UltimoMensaje = 4

            End If

            '[/CDT]
            
            Call SubirSkill(userIndex, eSkill.Ocultarse, False)

        End If
        
        .Counters.Ocultando = .Counters.Ocultando + 1

    End With
    
    Exit Sub

errHandler:
    Call LogError("Error en Sub DoOcultarse")

End Sub

Public Sub DoNavega(ByVal userIndex As Integer, _
                    ByRef Barco As ObjData, _
                    ByVal Slot As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 12/01/2020 (Recox)
'13/01/2010: ZaMa - El pirata pierde el ocultar si desequipa barca.
'16/09/2010: ZaMa - Ahora siempre se va el invi para los clientes al equipar la barca (Evita cortes de cabeza).
'10/12/2010: Pato - Limpio las variables del inventario que hacen referencia a la barca, sino el pirata que la ultima barca que equipo era el galeon no explotaba(Y capaz no la tenia equipada :P).
'12/01/2020: Recox - Se refactorizo un poco para reutilizar con monturas .
'***************************************************

    Dim ModNave As Single
    
    With UserList(userIndex)
        If .flags.Equitando = 1 Then
            Call WriteConsoleMsg(userIndex, "No puedes navegar mientras estas en tu montura!!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        '¿Es una montura acuatica? Pedimo Skills en Equitacion
        If Barco.MontTipo = 1 Then
            If UserList(userIndex).Stats.UserSkills(Equitacion) < Barco.MinSkill Then
                Call WriteConsoleMsg(userIndex, "Para usar esta montura necesitas " & Barco.MinSkill & " puntos en equitación.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        
        Else
        
            ModNave = ModNavegacion(.clase, userIndex)
            
            If .Stats.UserSkills(eSkill.Navegacion) / ModNave < Barco.MinSkill Then
                Call WriteConsoleMsg(userIndex, "No tienes suficientes conocimientos para usar este barco.", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(userIndex, "Para usar este barco necesitas " & Barco.MinSkill * ModNave & " puntos en navegacion.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
    
            End If
        
        End If
        
        ' No estaba navegando
        If .flags.Navegando = 0 Then
            .Invent.BarcoObjIndex = .Invent.Object(Slot).ObjIndex
            .Invent.BarcoSlot = Slot
            
            .Char.Head = 0
            
            ' No esta muerto
            If .flags.Muerto = 0 Then
                Call ToggleBoatBody(userIndex)
                Call SetVisibleStateForUserAfterNavigateOrEquitate(userIndex)
                
            ' Esta muerto
            Else
                .Char.body = iFragataFantasmal
                .Char.ShieldAnim = NingunEscudo
                .Char.WeaponAnim = NingunArma
                .Char.CascoAnim = NingunCasco
                .Char.AuraAnim = NingunAura
                .Char.AuraColor = NingunAura
                
            End If
            
            ' Comienza a navegar
            .flags.Navegando = 1
        
        ' Estaba navegando
        Else
            .Invent.BarcoObjIndex = 0
            .Invent.BarcoSlot = 0
        
            ' No esta muerto
            If .flags.Muerto = 0 Then
                .Char.Head = .OrigChar.Head
                
                Call SetEquipmentOnCharAfterNavigateOrEquitate(userIndex)
                
                ' Al dejar de navegar, si estaba invisible actualizo los clientes
                If .flags.invisible = 1 Then
                    Call SetInvisible(userIndex, .Char.CharIndex, True)
                End If
                
            ' Esta muerto
            Else
                .Char.body = iCuerpoMuerto
                .Char.Head = iCabezaMuerto
                .Char.ShieldAnim = NingunEscudo
                .Char.WeaponAnim = NingunArma
                .Char.CascoAnim = NingunCasco
                .Char.AuraAnim = NingunAura
                .Char.AuraColor = NingunAura

            End If
            
            ' Termina de navegar
            .flags.Navegando = 0

        End If
        
        ' Actualizo clientes
        Call ChangeUserChar(userIndex, .Char.body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraAnim, .Char.AuraColor)

    End With
    
    Call WriteNavigateToggle(userIndex)

End Sub

Public Sub FundirMineral(ByVal userIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo errHandler

    With UserList(userIndex)

        If .flags.TargetObjInvIndex > 0 Then
           
            If ObjData(.flags.TargetObjInvIndex).OBJType = eOBJType.otMinerales And ObjData(.flags.TargetObjInvIndex).MinSkill <= .Stats.UserSkills(eSkill.Mineria) / ModFundicion(.clase) Then
                Call DoLingotes(userIndex)
            Else
                Call WriteConsoleMsg(userIndex, "No tienes conocimientos de mineria suficientes para trabajar este mineral.", FontTypeNames.FONTTYPE_INFO)

            End If
        
        End If

    End With

    Exit Sub

errHandler:
    Call LogError("Error en FundirMineral. Error " & Err.Number & " : " & Err.description)

End Sub

Function TieneObjetos(ByVal ItemIndex As Integer, _
                      ByVal cant As Long, _
                      ByVal userIndex As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: 10/07/2010
    '10/07/2010: ZaMa - Ahora cant es long para evitar un overflow.
    '***************************************************

    Dim i     As Integer

    Dim Total As Long

    For i = 1 To UserList(userIndex).CurrentInventorySlots

        If UserList(userIndex).Invent.Object(i).ObjIndex = ItemIndex Then
            Total = Total + UserList(userIndex).Invent.Object(i).Amount

        End If

    Next i
    
    If cant <= Total Then
        TieneObjetos = True
        Exit Function

    End If
        
End Function

Public Sub QuitarObjetos(ByVal ItemIndex As Integer, _
                         ByVal cant As Integer, _
                         ByVal userIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: 05/08/09
    '05/08/09: Pato - Cambie la funcion a procedimiento ya que se usa como procedimiento siempre, y fixie el bug 2788199
    '***************************************************

    Dim i As Integer

    For i = 1 To UserList(userIndex).CurrentInventorySlots

        With UserList(userIndex).Invent.Object(i)

            If .ObjIndex = ItemIndex Then
                If .Amount <= cant And .Equipped = 1 Then Call Desequipar(userIndex, i)
                
                .Amount = .Amount - cant

                If .Amount <= 0 Then
                    cant = Abs(.Amount)
                    UserList(userIndex).Invent.NroItems = UserList(userIndex).Invent.NroItems - 1
                    .Amount = 0
                    .ObjIndex = 0
                Else
                    cant = 0

                End If
                
                Call UpdateUserInv(False, userIndex, i)
                
                If cant = 0 Then Exit Sub

            End If

        End With

    Next i

End Sub

Sub HerreroQuitarMateriales(ByVal userIndex As Integer, _
                            ByVal ItemIndex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: 16/11/2009
    '16/11/2009: ZaMa - Ahora considera la cantidad de items a construir
    '***************************************************
    With ObjData(ItemIndex)

        If .LingH > 0 Then Call QuitarObjetos(LingoteHierro, .LingH, userIndex)
        If .LingP > 0 Then Call QuitarObjetos(LingotePlata, .LingP, userIndex)
        If .LingO > 0 Then Call QuitarObjetos(LingoteOro, .LingO, userIndex)

    End With

End Sub

Sub CarpinteroQuitarMateriales(ByVal userIndex As Integer, _
                               ByVal ItemIndex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: 16/11/2009
    '16/11/2009: ZaMa - Ahora quita tambien madera elfica
    '***************************************************
    With ObjData(ItemIndex)

        If .Madera > 0 Then Call QuitarObjetos(Lena, .Madera, userIndex)
        If .MaderaElfica > 0 Then Call QuitarObjetos(LenaElfica, .MaderaElfica, userIndex)

    End With

End Sub

Function CarpinteroTieneMateriales(ByVal userIndex As Integer, _
                                   ByVal ItemIndex As Integer, _
                                   Optional ByVal ShowMsg As Boolean = False) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: 16/11/2009
    '16/11/2009: ZaMa - Agregada validacion a madera elfica.
    '16/11/2009: ZaMa - Ahora considera la cantidad de items a construir
    '***************************************************
    
    With ObjData(ItemIndex)

        If .Madera > 0 Then
            If Not TieneObjetos(Lena, .Madera, userIndex) Then
                If ShowMsg Then Call WriteConsoleMsg(userIndex, "No tienes suficiente madera.", FontTypeNames.FONTTYPE_INFO)
                CarpinteroTieneMateriales = False
                Exit Function

            End If

        End If
        
        If .MaderaElfica > 0 Then
            If Not TieneObjetos(LenaElfica, .MaderaElfica, userIndex) Then
                If ShowMsg Then Call WriteConsoleMsg(userIndex, "No tienes suficiente madera elfica.", FontTypeNames.FONTTYPE_INFO)
                CarpinteroTieneMateriales = False
                Exit Function

            End If

        End If
    
    End With

    CarpinteroTieneMateriales = True

End Function
 
Function HerreroTieneMateriales(ByVal userIndex As Integer, _
                                ByVal ItemIndex As Integer) As Boolean

    '***************************************************
    'Author: Unknown
    'Last Modification: 16/11/2009
    '16/11/2009: ZaMa - Agregada validacion a madera elfica.
    '***************************************************
    With ObjData(ItemIndex)

        If .LingH > 0 Then
            If Not TieneObjetos(LingoteHierro, .LingH, userIndex) Then
                Call WriteConsoleMsg(userIndex, "No tienes suficientes lingotes de hierro.", FontTypeNames.FONTTYPE_INFO)
                HerreroTieneMateriales = False
                Exit Function

            End If

        End If

        If .LingP > 0 Then
            If Not TieneObjetos(LingotePlata, .LingP, userIndex) Then
                Call WriteConsoleMsg(userIndex, "No tienes suficientes lingotes de plata.", FontTypeNames.FONTTYPE_INFO)
                HerreroTieneMateriales = False
                Exit Function

            End If

        End If

        If .LingO > 0 Then
            If Not TieneObjetos(LingoteOro, .LingO, userIndex) Then
                Call WriteConsoleMsg(userIndex, "No tienes suficientes lingotes de oro.", FontTypeNames.FONTTYPE_INFO)
                HerreroTieneMateriales = False
                Exit Function

            End If

        End If

    End With

    HerreroTieneMateriales = True

End Function

Public Function PuedeConstruirItemHerrero(ByVal userIndex As Integer, _
                               ByVal ItemIndex As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: 24/08/2009
    '24/08/2008: ZaMa - Validates if the player has the required skill
    '16/11/2009: ZaMa - Validates if the player has the required amount of materials, depending on the number of items to make
    '***************************************************
    PuedeConstruirItemHerrero = HerreroTieneMateriales(userIndex, ItemIndex) And Round(UserList(userIndex).Stats.UserSkills(eSkill.Herreria) / ModHerreriA(UserList(userIndex).clase), 0) >= ObjData(ItemIndex).SkHerreria

End Function

Public Function PuedeConstruirHerreria(ByVal ItemIndex As Integer) As Boolean

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    Dim i As Long

    For i = 1 To UBound(ArmasHerrero)

        If ArmasHerrero(i) = ItemIndex Then
            PuedeConstruirHerreria = True
            Exit Function

        End If

    Next i

    For i = 1 To UBound(ArmadurasHerrero)

        If ArmadurasHerrero(i) = ItemIndex Then
            PuedeConstruirHerreria = True
            Exit Function

        End If

    Next i

    PuedeConstruirHerreria = False

End Function

Public Sub HerreroConstruirItem(ByVal userIndex As Integer, ByVal ItemIndex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: 30/05/2010
    '16/11/2009: ZaMa - Implementado nuevo sistema de construccion de items.
    '22/05/2010: ZaMa - Los caos ya no suben plebe al trabajar.
    '30/05/2010: ZaMa - Los pks no suben plebe al trabajar.
    '***************************************************

    Dim TieneMateriales As Boolean

    Dim OtroUserIndex   As Integer

    With UserList(userIndex)

        If .flags.Comerciando Then
            OtroUserIndex = .ComUsu.DestUsu
            
            If OtroUserIndex > 0 And OtroUserIndex <= MaxUsers Then
                Call WriteConsoleMsg(userIndex, "Comercio cancelado, no puedes comerciar mientras trabajas!!", FontTypeNames.FONTTYPE_TALK)
                Call WriteConsoleMsg(OtroUserIndex, "Comercio cancelado por el otro usuario!!", FontTypeNames.FONTTYPE_TALK)
            
                Call LimpiarComercioSeguro(userIndex)

            End If

        End If
    
        If PuedeConstruirHerreria(ItemIndex) Then
        
            'Sacamos energia
            'Chequeamos que tenga los puntos antes de sacarselos
            If .Stats.MinSta >= GASTO_ENERGIA Then
                .Stats.MinSta = .Stats.MinSta - GASTO_ENERGIA
                Call WriteUpdateSta(userIndex)
            Else
                Call WriteConsoleMsg(userIndex, "No tienes suficiente energia.", FontTypeNames.FONTTYPE_INFO)
                Call DejardeTrabajar(userIndex) 'Paramos el macro
                Exit Sub

            End If

        
            Call HerreroQuitarMateriales(userIndex, ItemIndex)
            ' AGREGAR FX
        
            'Mensajes de exito
            Select Case ObjData(ItemIndex).OBJType
                Case eOBJType.otWeapon
                    Call WriteConsoleMsg(userIndex, "Has construido el arma!.", FontTypeNames.FONTTYPE_INFO)
                    
                Case eOBJType.otEscudo
                    Call WriteConsoleMsg(userIndex, "Has construido el escudo!.", FontTypeNames.FONTTYPE_INFO)
                    
                Case eOBJType.otCasco
                    Call WriteConsoleMsg(userIndex, "Has construido el casco!.", FontTypeNames.FONTTYPE_INFO)
                    
                Case eOBJType.otArmadura
                    Call WriteConsoleMsg(userIndex, "Has construido la armadura!.", FontTypeNames.FONTTYPE_INFO)
            End Select
        
            Dim MiObj As obj
        
            MiObj.Amount = 1
            MiObj.ObjIndex = ItemIndex

            If Not MeterItemEnInventario(userIndex, MiObj) Then
                Call TirarItemAlPiso(.Pos, MiObj)

            End If
        
            'Log de construccion de Items. Pablo (ToxicWaste) 10/09/07
            If ObjData(MiObj.ObjIndex).Log = 1 Then
                Call LogDesarrollo(.Name & " ha construido " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).Name)

            End If
        
            Call SubirSkill(userIndex, eSkill.Herreria, True)
            Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessagePlayWave(SND_TRABAJO_HERRERO, .Pos.X, .Pos.Y))
        
            If Not criminal(userIndex) Then
                .Reputacion.PlebeRep = .Reputacion.PlebeRep + vlProleta

                If .Reputacion.PlebeRep > MAXREP Then .Reputacion.PlebeRep = MAXREP

            End If
        
            .Counters.Trabajando = .Counters.Trabajando + 1

        End If

    End With

End Sub

Public Function PuedeConstruirCarpintero(ByVal ItemIndex As Integer) As Boolean

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    Dim i As Long

    For i = 1 To UBound(ObjCarpintero)

        If ObjCarpintero(i) = ItemIndex Then
            PuedeConstruirCarpintero = True
            Exit Function

        End If

    Next i

    PuedeConstruirCarpintero = False

End Function

Public Sub CarpinteroConstruirItem(ByVal userIndex As Integer, ByVal ItemIndex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: 28/05/2010
    '24/08/2008: ZaMa - Validates if the player has the required skill
    '16/11/2009: ZaMa - Implementado nuevo sistema de construccion de items
    '22/05/2010: ZaMa - Los caos ya no suben plebe al trabajar.
    '28/05/2010: ZaMa - Los pks no suben plebe al trabajar.
    '***************************************************
    On Error GoTo errHandler

    Dim TieneMateriales As Boolean

    Dim WeaponIndex     As Integer

    Dim OtroUserIndex   As Integer
    
    With UserList(userIndex)

        If .flags.Comerciando Then
            OtroUserIndex = .ComUsu.DestUsu
                
            If OtroUserIndex > 0 And OtroUserIndex <= MaxUsers Then
                Call WriteConsoleMsg(userIndex, "Comercio cancelado, no puedes comerciar mientras trabajas!!", FontTypeNames.FONTTYPE_TALK)
                Call WriteConsoleMsg(OtroUserIndex, "Comercio cancelado por el otro usuario!!", FontTypeNames.FONTTYPE_TALK)
                
                Call LimpiarComercioSeguro(userIndex)

            End If

        End If
        
        WeaponIndex = .Invent.WeaponEqpObjIndex
    
        If WeaponIndex <> SERRUCHO_CARPINTERO Then
            Call WriteConsoleMsg(userIndex, "Debes tener equipado el serrucho para trabajar.", FontTypeNames.FONTTYPE_INFO)
            Call DejardeTrabajar(userIndex) 'Paramos el macro
            Exit Sub

        End If
    
        If Round(.Stats.UserSkills(eSkill.Carpinteria) \ ModCarpinteria(.clase), 0) >= ObjData(ItemIndex).SkCarpinteria And PuedeConstruirCarpintero(ItemIndex) Then
           
            'Sacamos energia
            'Chequeamos que tenga los puntos antes de sacarselos
            If .Stats.MinSta >= GASTO_ENERGIA Then
                .Stats.MinSta = .Stats.MinSta - GASTO_ENERGIA
                Call WriteUpdateSta(userIndex)
            Else
                Call WriteConsoleMsg(userIndex, "No tienes suficiente energia.", FontTypeNames.FONTTYPE_INFO)
                Call DejardeTrabajar(userIndex) 'Paramos el macro
                Exit Sub

            End If
            
            Call CarpinteroQuitarMateriales(userIndex, ItemIndex)
            Call WriteConsoleMsg(userIndex, "Has construido el objeto!.", FontTypeNames.FONTTYPE_INFO)
            
            Dim MiObj As obj

            MiObj.Amount = 1
            MiObj.ObjIndex = ItemIndex

            If Not MeterItemEnInventario(userIndex, MiObj) Then
                Call TirarItemAlPiso(.Pos, MiObj)

            End If
            
            'Log de construccion de Items. Pablo (ToxicWaste) 10/09/07
            If ObjData(MiObj.ObjIndex).Log = 1 Then
                Call LogDesarrollo(.Name & " ha construido " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).Name)

            End If
            
            Call SubirSkill(userIndex, eSkill.Carpinteria, True)
            Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessagePlayWave(SND_TRABAJO_CARPINTERO, .Pos.X, .Pos.Y))
            
            If Not criminal(userIndex) Then
                .Reputacion.PlebeRep = .Reputacion.PlebeRep + vlProleta

                If .Reputacion.PlebeRep > MAXREP Then .Reputacion.PlebeRep = MAXREP

            End If
            
            .Counters.Trabajando = .Counters.Trabajando + 1

        End If

    End With
    
    Exit Sub
errHandler:
    Call LogError("Error en CarpinteroConstruirItem. Error " & Err.Number & " : " & Err.description & ". UserIndex:" & userIndex & ". ItemIndex:" & ItemIndex)

End Sub

Public Sub ArtesanoConstruirItem(ByVal userIndex As Integer, ByVal Item As Integer)
    Dim ArtesanoObj As ObjData
    ArtesanoObj = ObjData(ObjArtesano(Item))

    Dim NPCIndex As Integer
    NPCIndex = UserList(userIndex).flags.TargetNPC

    ' Revisamos si tiene las monedas para la comision
    If UserList(userIndex).Stats.Gld < ArtesaniaCosto Then
        Call WriteChatOverHead(userIndex, "No tienes suficientes monedas de oro para pagarme!", Npclist(NPCIndex).Char.CharIndex, vbWhite)
        Exit Sub
    End If

    ' Revisamos si tiene los materiales necesarios
    Dim i As Integer
    For i = 1 To UBound(ArtesanoObj.ItemCrafteo)

        With ArtesanoObj.ItemCrafteo(i)

            If Not TieneObjetos(.ObjIndex, .Amount, userIndex) Then
                Call WriteChatOverHead(userIndex, "No tienes los materiales necesarios!", Npclist(NPCIndex).Char.CharIndex, vbWhite)
                Exit Sub
            End If

        End With

    Next i

    ' Le sacamos el oro
    UserList(userIndex).Stats.Gld = UserList(userIndex).Stats.Gld - ArtesaniaCosto
    Call WriteUpdateGold(userIndex)
    
    Call WriteConsoleMsg(userIndex, "Le has pagado " & Format$(ArtesaniaCosto, "##,##") & " monedas de oro al artesano.", FontTypeNames.FONTTYPE_INFO)

    ' Le sacamos los materiales
    For i = 1 To UBound(ArtesanoObj.ItemCrafteo)

        With ArtesanoObj.ItemCrafteo(i)

            Call QuitarObjetos(.ObjIndex, .Amount, userIndex)

        End With

    Next i

    Dim ObjetoCreado As obj
    ObjetoCreado.ObjIndex = ObjArtesano(Item)
    ObjetoCreado.Amount = 1

    ' Metemos el item en el inventario o lo tiramos al piso
    If Not MeterItemEnInventario(userIndex, ObjetoCreado) Then
        Call TirarItemAlPiso(UserList(userIndex).Pos, ObjetoCreado)
    End If

    Call WriteChatOverHead(userIndex, "Aqui tienes tu " & ArtesanoObj.Name & ". Vuelve pronto!", Npclist(NPCIndex).Char.CharIndex, vbWhite)

End Sub

Private Function MineralesParaLingote(ByVal Lingote As iMinerales) As Integer

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    Select Case Lingote

        Case iMinerales.HierroCrudo
            MineralesParaLingote = 14

        Case iMinerales.PlataCruda
            MineralesParaLingote = 20

        Case iMinerales.OroCrudo
            MineralesParaLingote = 35

        Case Else
            MineralesParaLingote = 10000

    End Select

End Function

Public Sub DoLingotes(ByVal userIndex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: 16/11/2009
    '16/11/2009: ZaMa - Implementado nuevo sistema de construccion de items
    '***************************************************
    '    Call LogTarea("Sub DoLingotes")
    Dim Slot           As Integer

    Dim obji           As Integer

    Dim CantidadItems  As Integer

    Dim TieneMinerales As Boolean

    Dim OtroUserIndex  As Integer
    
    With UserList(userIndex)

        If .flags.Comerciando Then
            OtroUserIndex = .ComUsu.DestUsu
                
            If OtroUserIndex > 0 And OtroUserIndex <= MaxUsers Then
                Call WriteConsoleMsg(userIndex, "Comercio cancelado, no puedes comerciar mientras trabajas!!", FontTypeNames.FONTTYPE_TALK)
                Call WriteConsoleMsg(OtroUserIndex, "Comercio cancelado por el otro usuario!!", FontTypeNames.FONTTYPE_TALK)
                
                Call LimpiarComercioSeguro(userIndex)

            End If

        End If
        
        CantidadItems = MaximoInt(1, CInt((.Stats.ELV - 4) / 5))

        Slot = .flags.TargetObjInvSlot
        obji = .Invent.Object(Slot).ObjIndex
        
        While CantidadItems > 0 And Not TieneMinerales

            If .Invent.Object(Slot).Amount >= MineralesParaLingote(obji) * CantidadItems Then
                TieneMinerales = True
            Else
                CantidadItems = CantidadItems - 1

            End If

        Wend
        
        If Not TieneMinerales Or ObjData(obji).OBJType <> eOBJType.otMinerales Then
            Call WriteConsoleMsg(userIndex, "No tienes suficientes minerales para hacer un lingote.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
        
        .Invent.Object(Slot).Amount = .Invent.Object(Slot).Amount - MineralesParaLingote(obji) * CantidadItems

        If .Invent.Object(Slot).Amount < 1 Then
            .Invent.Object(Slot).Amount = 0
            .Invent.Object(Slot).ObjIndex = 0

        End If
        
        Dim MiObj As obj

        MiObj.Amount = CantidadItems
        MiObj.ObjIndex = ObjData(.flags.TargetObjInvIndex).LingoteIndex

        If Not MeterItemEnInventario(userIndex, MiObj) Then
            Call TirarItemAlPiso(.Pos, MiObj)

        End If
        
        Call UpdateUserInv(False, userIndex, Slot)
        Call WriteConsoleMsg(userIndex, "Has obtenido " & CantidadItems & " lingote" & IIf(CantidadItems = 1, "", "s") & "!", FontTypeNames.FONTTYPE_INFO)
    
        .Counters.Trabajando = .Counters.Trabajando + 1

    End With

End Sub

Public Sub DoFundir(ByVal userIndex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: 03/06/2010
    '03/06/2010 - Pato: Si es el ultimo item a fundir y esta equipado lo desequipamos.
    '11/03/2010 - ZaMa: Reemplazo division por producto para uan mejor performanse.
    '***************************************************
    Dim i             As Integer

    Dim Num           As Integer

    Dim Slot          As Byte

    Dim Lingotes(2)   As Integer

    Dim OtroUserIndex As Integer

    With UserList(userIndex)

        If .flags.Comerciando Then
            OtroUserIndex = .ComUsu.DestUsu
                
            If OtroUserIndex > 0 And OtroUserIndex <= MaxUsers Then
                Call WriteConsoleMsg(userIndex, "Comercio cancelado, no puedes comerciar mientras trabajas!!", FontTypeNames.FONTTYPE_TALK)
                Call WriteConsoleMsg(OtroUserIndex, "Comercio cancelado por el otro usuario!!", FontTypeNames.FONTTYPE_TALK)
                
                Call LimpiarComercioSeguro(userIndex)

            End If

        End If
        
        Slot = .flags.TargetObjInvSlot
        
        With .Invent.Object(Slot)
            .Amount = .Amount - 1
            
            If .Amount < 1 Then
                If .Equipped = 1 Then Call Desequipar(userIndex, Slot)
                
                .Amount = 0
                .ObjIndex = 0

            End If

        End With
        
        Num = RandomNumber(10, 25)
        
        Lingotes(0) = (ObjData(.flags.TargetObjInvIndex).LingH * Num) * 0.01
        Lingotes(1) = (ObjData(.flags.TargetObjInvIndex).LingP * Num) * 0.01
        Lingotes(2) = (ObjData(.flags.TargetObjInvIndex).LingO * Num) * 0.01
    
        Dim MiObj(2) As obj
        
        For i = 0 To 2
            MiObj(i).Amount = Lingotes(i)
            MiObj(i).ObjIndex = LingoteHierro + i 'Una gran negrada pero practica
            
            If MiObj(i).Amount > 0 Then
                If Not MeterItemEnInventario(userIndex, MiObj(i)) Then
                    Call TirarItemAlPiso(.Pos, MiObj(i))

                End If

            End If

        Next i
        
        Call UpdateUserInv(False, userIndex, Slot)
        Call WriteConsoleMsg(userIndex, "Has obtenido el " & Num & "% de los lingotes utilizados para la construccion del objeto!", FontTypeNames.FONTTYPE_INFO)
    
        .Counters.Trabajando = .Counters.Trabajando + 1

    End With

End Sub

Function ModNavegacion(ByVal clase As eClass, ByVal userIndex As Integer) As Single

    '***************************************************
    'Autor: Unknown (orginal version)
    'Last Modification: 27/11/2009
    '12/04/2010: ZaMa - Arreglo modificador de pescador, para que navegue con 60 skills.
    '***************************************************
    Select Case clase

        Case eClass.Pirat
            ModNavegacion = 1

        Case Else
            ModNavegacion = 2

    End Select

End Function

Function ModFundicion(ByVal clase As eClass) As Single
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    ModFundicion = 3

End Function

Function ModCarpinteria(ByVal clase As eClass) As Integer
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    ModCarpinteria = 3


End Function

Function ModHerreriA(ByVal clase As eClass) As Single

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    
    ModHerreriA = 4

End Function

Function ModDomar(ByVal clase As eClass) As Integer

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    Select Case clase

        Case eClass.Druid
            ModDomar = 6

        Case eClass.Hunter
            ModDomar = 6

        Case eClass.Cleric
            ModDomar = 7

        Case Else
            ModDomar = 10

    End Select

End Function

Function FreeMascotaIndex(ByVal userIndex As Integer) As Integer

    '***************************************************
    'Author: Unknown
    'Last Modification: 02/03/09
    '02/03/09: ZaMa - Busca un indice libre de mascotas, revisando los types y no los indices de los npcs
    '***************************************************
    Dim j As Integer

    For j = 1 To MAXMASCOTAS

        If UserList(userIndex).MascotasType(j) = 0 Then
            FreeMascotaIndex = j
            Exit Function

        End If

    Next j

End Function

Sub DoDomar(ByVal userIndex As Integer, ByVal NPCIndex As Integer)
    '***************************************************
    'Author: Nacho (Integer)
    'Last Modification: 01/05/2010
    '12/15/2008: ZaMa - Limits the number of the same type of pet to 2.
    '02/03/2009: ZaMa - Las criaturas domadas en zona segura, esperan afuera (desaparecen).
    '01/05/2010: ZaMa - Agrego bonificacion 11% para domar con flauta magica.
    '***************************************************

    On Error GoTo errHandler

    Dim puntosDomar      As Integer

    Dim puntosRequeridos As Integer

    Dim CanStay          As Boolean

    Dim petType          As Integer

    Dim NroPets          As Integer
    
    If Npclist(NPCIndex).MaestroUser = userIndex Then
        Call WriteConsoleMsg(userIndex, "Ya domaste a esa criatura.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub

    End If

    With UserList(userIndex)

        If .NroMascotas < MAXMASCOTAS Then
            
            If Npclist(NPCIndex).MaestroNpc > 0 Or Npclist(NPCIndex).MaestroUser > 0 Then
                Call WriteConsoleMsg(userIndex, "La criatura ya tiene amo.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
            
            If Not PuedeDomarMascota(userIndex, NPCIndex) Then
                Call WriteConsoleMsg(userIndex, "No puedes domar mas de dos criaturas del mismo tipo.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
            
            puntosDomar = CInt(.Stats.UserAtributos(eAtributos.Carisma)) * CInt(.Stats.UserSkills(eSkill.Domar))
            
            ' 20% de bonificacion
            If .Invent.AnilloEqpObjIndex = FLAUTAELFICA Then
                puntosRequeridos = Npclist(NPCIndex).flags.Domable * 0.8
            
                ' 11% de bonificacion
            ElseIf .Invent.AnilloEqpObjIndex = FLAUTAMAGICA Then
                puntosRequeridos = Npclist(NPCIndex).flags.Domable * 0.89
                
            Else
                puntosRequeridos = Npclist(NPCIndex).flags.Domable

            End If
            
            If puntosRequeridos <= puntosDomar And RandomNumber(1, 5) = 1 Then

                Dim index As Integer

                .NroMascotas = .NroMascotas + 1
                index = FreeMascotaIndex(userIndex)
                .MascotasIndex(index) = NPCIndex
                .MascotasType(index) = Npclist(NPCIndex).Numero
                
                Npclist(NPCIndex).MaestroUser = userIndex
                
                Call FollowAmo(NPCIndex)
                Call ReSpawnNpc(Npclist(NPCIndex))
                
                Call WriteConsoleMsg(userIndex, "La criatura te ha aceptado como su amo.", FontTypeNames.FONTTYPE_INFO)
                
                ' Es zona segura?
                CanStay = (MapInfo(.Pos.Map).Pk = True)
                
                If Not CanStay Then
                    petType = Npclist(NPCIndex).Numero
                    NroPets = .NroMascotas
                    
                    Call QuitarNPC(NPCIndex)
                    
                    .MascotasType(index) = petType
                    .NroMascotas = NroPets
                    
                    Call WriteConsoleMsg(userIndex, "No se permiten mascotas en zona segura. estas te esperaran afuera.", FontTypeNames.FONTTYPE_INFO)

                End If
                
                Call SubirSkill(userIndex, eSkill.Domar, True)
        
            Else

                If Not .flags.UltimoMensaje = 5 Then
                    Call WriteConsoleMsg(userIndex, "No has logrado domar la criatura.", FontTypeNames.FONTTYPE_INFO)
                    .flags.UltimoMensaje = 5

                End If
                
                Call SubirSkill(userIndex, eSkill.Domar, False)

            End If

        Else
            Call WriteConsoleMsg(userIndex, "No puedes controlar mas criaturas.", FontTypeNames.FONTTYPE_INFO)

        End If

    End With
    
    Exit Sub

errHandler:
    Call LogError("Error en DoDomar. Error " & Err.Number & " : " & Err.description)

End Sub

''
' Checks if the user can tames a pet.
'
' @param integer userIndex The user id from who wants tame the pet.
' @param integer NPCindex The index of the npc to tome.
' @return boolean True if can, false if not.
Private Function PuedeDomarMascota(ByVal userIndex As Integer, _
                                   ByVal NPCIndex As Integer) As Boolean

    '***************************************************
    'Author: ZaMa
    'This function checks how many NPCs of the same type have
    'been tamed by the user.
    'Returns True if that amount is less than two.
    '***************************************************
    Dim i           As Long

    Dim numMascotas As Long
    
    For i = 1 To MAXMASCOTAS

        If UserList(userIndex).MascotasType(i) = Npclist(NPCIndex).Numero Then
            numMascotas = numMascotas + 1

        End If

    Next i
    
    If numMascotas <= 1 Then PuedeDomarMascota = True
    
End Function

Sub DoAdminInvisible(ByVal userIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: 12/01/2010 (ZaMa)
    'Makes an admin invisible o visible.
    '13/07/2009: ZaMa - Now invisible admins' chars are erased from all clients, except from themselves.
    '12/01/2010: ZaMa - Los druidas pierden la inmunidad de ser atacados cuando pierden el efecto del mimetismo.
    '***************************************************
    
    Dim tempData As String
    
    With UserList(userIndex)

        If .flags.AdminInvisible = 0 Then

            ' Sacamos el mimetizmo
            If .flags.Mimetizado = 1 Then
                .Char.body = .CharMimetizado.body
                .Char.Head = .CharMimetizado.Head
                .Char.CascoAnim = .CharMimetizado.CascoAnim
                .Char.ShieldAnim = .CharMimetizado.ShieldAnim
                .Char.WeaponAnim = .CharMimetizado.WeaponAnim
                .Counters.Mimetismo = 0
                .flags.Mimetizado = 0
                ' Se fue el efecto del mimetismo, puede ser atacado por npcs
                .flags.Ignorado = False

            End If
            
            'Guardamos el antiguo body y head
            .flags.OldBody = .Char.body
            .flags.OldHead = .Char.Head
            
            .flags.AdminInvisible = 1
            .flags.invisible = 1
            .flags.Oculto = 1
            
            ' Solo el admin sabe que se hace invi
            tempData = PrepareMessageSetInvisible(.Char.CharIndex, True)
            Call UserList(userIndex).outgoingData.WriteASCIIStringFixed(tempData)
            
            'Le mandamos el mensaje para que borre el personaje a los clientes que esten cerca
            Call SendData(SendTarget.ToPCAreaButIndex, userIndex, PrepareMessageCharacterRemove(.Char.CharIndex))
            
        Else
            .flags.AdminInvisible = 0
            .flags.invisible = 0
            .flags.Oculto = 0
            .Counters.TiempoOculto = 0
            
            ' Solo el admin sabe que se hace visible
            tempData = PrepareMessageCharacterChange(.Char.body, .Char.Head, .Char.Heading, .Char.CharIndex, .Char.WeaponAnim, .Char.ShieldAnim, .Char.FX, .Char.loops, .Char.CascoAnim, .Char.AuraAnim, .Char.AuraColor)
            Call UserList(userIndex).outgoingData.WriteASCIIStringFixed(tempData)
            
            tempData = PrepareMessageSetInvisible(.Char.CharIndex, False)
            Call UserList(userIndex).outgoingData.WriteASCIIStringFixed(tempData)
             
            'Le mandamos el mensaje para crear el personaje a los clientes que esten cerca
            Call MakeUserChar(True, .Pos.Map, userIndex, .Pos.Map, .Pos.X, .Pos.Y, True)

        End If

    End With
    
End Sub

Sub TratarDeHacerFogata(ByVal Map As Integer, _
                        ByVal X As Integer, _
                        ByVal Y As Integer, _
                        ByVal userIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim Suerte    As Byte

    Dim exito     As Byte

    Dim obj       As obj

    Dim posMadera As WorldPos

    If Not LegalPos(Map, X, Y) Then Exit Sub

    With posMadera
        .Map = Map
        .X = X
        .Y = Y

    End With

    If MapData(Map, X, Y).ObjInfo.ObjIndex <> 58 Then
        Call WriteConsoleMsg(userIndex, "Necesitas clickear sobre lena para hacer ramitas.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub

    End If

    If Distancia(posMadera, UserList(userIndex).Pos) > 2 Then
        Call WriteConsoleMsg(userIndex, "Estas demasiado lejos para prender la fogata.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub

    End If

    If UserList(userIndex).flags.Muerto = 1 Then
        Call WriteConsoleMsg(userIndex, "No puedes hacer fogatas estando muerto.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub

    End If

    If MapData(Map, X, Y).ObjInfo.Amount < 3 Then
        Call WriteConsoleMsg(userIndex, "Necesitas por lo menos tres troncos para hacer una fogata.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub

    End If

    Dim SupervivenciaSkill As Byte

    SupervivenciaSkill = UserList(userIndex).Stats.UserSkills(eSkill.Supervivencia)

    If SupervivenciaSkill < 6 Then
        Suerte = 3
    ElseIf SupervivenciaSkill <= 34 Then
        Suerte = 2
    Else
        Suerte = 1

    End If

    exito = RandomNumber(1, Suerte)

    If exito = 1 Then
        obj.ObjIndex = FOGATA_APAG
        obj.Amount = MapData(Map, X, Y).ObjInfo.Amount \ 3
    
        Call WriteConsoleMsg(userIndex, "Has hecho " & obj.Amount & " fogatas.", FontTypeNames.FONTTYPE_INFO)
    
        Call MakeObj(obj, Map, X, Y)
    
        'Seteamos la fogata como el nuevo TargetObj del user
        UserList(userIndex).flags.TargetObj = FOGATA_APAG
    
        Call SubirSkill(userIndex, eSkill.Supervivencia, True)
    Else

        '[CDT 17-02-2004]
        If Not UserList(userIndex).flags.UltimoMensaje = 10 Then
            Call WriteConsoleMsg(userIndex, "No has podido hacer la fogata.", FontTypeNames.FONTTYPE_INFO)
            UserList(userIndex).flags.UltimoMensaje = 10

        End If

        '[/CDT]
    
        Call SubirSkill(userIndex, eSkill.Supervivencia, False)

    End If

End Sub

Public Sub DoPescar(ByVal userIndex As Integer, ByVal Red As Boolean)

    '***************************************************
    'Author: Unknown
    'Last Modification: 26/10/2018
    '26/10/2018: CHOTS - Multiplicador de oficios
    '***************************************************
    On Error GoTo errHandler

    Dim iSkill        As Integer

    Dim Suerte        As Integer

    Dim res           As Integer

    Dim MAXITEMS      As Integer

    Dim CantidadItems As Integer

    With UserList(userIndex)
    
        Call QuitarSta(userIndex, EsfuerzoExtraer)

        iSkill = .Stats.UserSkills(eSkill.pesca)
        
        ' m = (60-11)/(1-10)
        ' y = mx - m*10 + 11
        
        Suerte = Int(-0.00125 * iSkill * iSkill - 0.3 * iSkill + 49)

        If Suerte > 0 Then
            res = RandomNumber(1, Suerte)
            
            If res <= DificultadExtraer Then
            
                Dim MiObj As obj
                
                MAXITEMS = MaxItemsExtraibles(.Stats.ELV)
                CantidadItems = RandomNumber(1, MAXITEMS)
                    
                CantidadItems = CantidadItems * OficioMultiplier
                
                MiObj.Amount = CantidadItems
                
                If Red Then
                    MiObj.ObjIndex = ListaPeces(RandomNumber(1, NUM_PECES))
                Else
                    MiObj.ObjIndex = Pescado
                End If
                
                If Not MeterItemEnInventario(userIndex, MiObj) Then
                    Call TirarItemAlPiso(.Pos, MiObj)

                End If
                
                Call WriteConsoleMsg(userIndex, "Has pescado algunos peces!", FontTypeNames.FONTTYPE_INFO)
                
                Call SubirSkill(userIndex, eSkill.pesca, True)
            Else

                If Not .flags.UltimoMensaje = 6 Then
                    Call WriteConsoleMsg(userIndex, "No has pescado nada!", FontTypeNames.FONTTYPE_INFO)
                    .flags.UltimoMensaje = 6

                End If
                
                Call SubirSkill(userIndex, eSkill.pesca, False)

            End If

        End If
        
        .Reputacion.PlebeRep = .Reputacion.PlebeRep + vlProleta

        If .Reputacion.PlebeRep > MAXREP Then .Reputacion.PlebeRep = MAXREP
        
        'Sonido
        Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessagePlayWave(SND_PESCAR, .Pos.X, .Pos.Y))
    
    End With
    
    Exit Sub

errHandler:
    Call LogError("Error en DoPescar Red: " & Red)

End Sub

''
' Try to steal an item / gold to another character
'
' @param LadrOnIndex Specifies reference to user that stoles
' @param VictimaIndex Specifies reference to user that is being stolen

Public Sub DoRobar(ByVal LadrOnIndex As Integer, ByVal VictimaIndex As Integer)
    '*************************************************
    'Author: Unknown
    'Last modified: 05/04/2010
    'Last Modification By: ZaMa
    '24/07/08: Marco - Now it calls to WriteUpdateGold(VictimaIndex and LadrOnIndex) when the thief stoles gold. (MarKoxX)
    '27/11/2009: ZaMa - Optimizacion de codigo.
    '18/12/2009: ZaMa - Los ladrones ciudas pueden robar a pks.
    '01/04/2010: ZaMa - Los ladrones pasan a robar oro acorde a su nivel.
    '05/04/2010: ZaMa - Los armadas no pueden robarle a ciudadanos jamas.
    '23/04/2010: ZaMa - No se puede robar mas sin energia.
    '23/04/2010: ZaMa - El alcance de robo pasa a ser de 1 tile.
    '*************************************************

    On Error GoTo errHandler

    Dim OtroUserIndex As Integer

    If Not MapInfo(UserList(VictimaIndex).Pos.Map).Pk Then Exit Sub
    
    If UserList(VictimaIndex).flags.EnConsulta Then
        Call WriteConsoleMsg(LadrOnIndex, "No puedes robar a usuarios en consulta!!!", FontTypeNames.FONTTYPE_INFO)
        Exit Sub

    End If
    
    With UserList(LadrOnIndex)
    
        If .flags.Seguro Then
            If Not criminal(VictimaIndex) Then
                Call WriteConsoleMsg(LadrOnIndex, "Debes quitarte el seguro para robarle a un ciudadano.", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub

            End If

        Else

            If .Faccion.ArmadaReal = 1 Then
                If Not criminal(VictimaIndex) Then
                    Call WriteConsoleMsg(LadrOnIndex, "Los miembros del ejercito real no tienen permitido robarle a ciudadanos.", FontTypeNames.FONTTYPE_FIGHT)
                    Exit Sub

                End If

            End If

        End If
        
        ' Caos robando a caos?
        If UserList(VictimaIndex).Faccion.FuerzasCaos = 1 And .Faccion.FuerzasCaos = 1 Then
            Call WriteConsoleMsg(LadrOnIndex, "No puedes robar a otros miembros de la legion oscura.", FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub

        End If
        
        If TriggerZonaPelea(LadrOnIndex, VictimaIndex) <> TRIGGER6_AUSENTE Then Exit Sub
        
        ' Tiene energia?
        If .Stats.MinSta < 15 Then
            If .Genero = eGenero.Hombre Then
                Call WriteConsoleMsg(LadrOnIndex, "Estas muy cansado para robar.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(LadrOnIndex, "Estas muy cansada para robar.", FontTypeNames.FONTTYPE_INFO)

            End If
            
            Exit Sub

        End If
        
        ' Quito energia
        Call QuitarSta(LadrOnIndex, 15)
        
        Dim GuantesHurto As Boolean
    
        If .Invent.AnilloEqpObjIndex = GUANTE_HURTO Then GuantesHurto = True
        
        If UserList(VictimaIndex).flags.Privilegios And PlayerType.User Then
            
            Dim Suerte     As Integer

            Dim res        As Integer

            Dim RobarSkill As Byte
            
            RobarSkill = .Stats.UserSkills(eSkill.Robar)
                
            If RobarSkill <= 10 Then
                Suerte = 35
            ElseIf RobarSkill <= 20 Then
                Suerte = 30
            ElseIf RobarSkill <= 30 Then
                Suerte = 28
            ElseIf RobarSkill <= 40 Then
                Suerte = 24
            ElseIf RobarSkill <= 50 Then
                Suerte = 22
            ElseIf RobarSkill <= 60 Then
                Suerte = 20
            ElseIf RobarSkill <= 70 Then
                Suerte = 18
            ElseIf RobarSkill <= 80 Then
                Suerte = 15
            ElseIf RobarSkill <= 90 Then
                Suerte = 10
            ElseIf RobarSkill < 100 Then
                Suerte = 7
            Else
                Suerte = 5

            End If
            
            res = RandomNumber(1, Suerte)
                
            If res < 3 Then 'Exito robo
                If UserList(VictimaIndex).flags.Comerciando Then
                    OtroUserIndex = UserList(VictimaIndex).ComUsu.DestUsu
                        
                    If OtroUserIndex > 0 And OtroUserIndex <= MaxUsers Then
                        Call WriteConsoleMsg(VictimaIndex, "Comercio cancelado, te estan robando!!", FontTypeNames.FONTTYPE_TALK)
                        Call WriteConsoleMsg(OtroUserIndex, "Comercio cancelado por el otro usuario!!", FontTypeNames.FONTTYPE_TALK)
                        
                        Call LimpiarComercioSeguro(VictimaIndex)

                    End If

                End If
               
                If (RandomNumber(1, 50) < 25) And (.clase = eClass.Thief) Then
                    If TieneObjetosRobables(VictimaIndex) Then
                        Call RobarObjeto(LadrOnIndex, VictimaIndex)
                    Else
                        Call WriteConsoleMsg(LadrOnIndex, UserList(VictimaIndex).Name & " no tiene objetos.", FontTypeNames.FONTTYPE_INFO)

                    End If

                Else 'Roba oro

                    If UserList(VictimaIndex).Stats.Gld > 0 Then

                        Dim n As Long
                        
                        If .clase = eClass.Thief Then

                            ' Si no tine puestos los guantes de hurto roba un 50% menos. Pablo (ToxicWaste)
                            If GuantesHurto Then
                                n = RandomNumber(.Stats.ELV * 50, .Stats.ELV * 100)
                            Else
                                n = RandomNumber(.Stats.ELV * 25, .Stats.ELV * 50)

                            End If

                        Else
                            n = RandomNumber(1, 100)

                        End If

                        If n > UserList(VictimaIndex).Stats.Gld Then n = UserList(VictimaIndex).Stats.Gld
                        UserList(VictimaIndex).Stats.Gld = UserList(VictimaIndex).Stats.Gld - n
                        
                        .Stats.Gld = .Stats.Gld + n

                        If .Stats.Gld > MAXORO Then .Stats.Gld = MAXORO
                        
                        Call WriteConsoleMsg(LadrOnIndex, "Le has robado " & n & " monedas de oro a " & UserList(VictimaIndex).Name, FontTypeNames.FONTTYPE_INFO)
                        Call WriteUpdateGold(LadrOnIndex) 'Le actualizamos la billetera al ladron
                        
                        Call WriteUpdateGold(VictimaIndex) 'Le actualizamos la billetera a la victima
                    Else
                        Call WriteConsoleMsg(LadrOnIndex, UserList(VictimaIndex).Name & " no tiene oro.", FontTypeNames.FONTTYPE_INFO)

                    End If

                End If
                
                Call SubirSkill(LadrOnIndex, eSkill.Robar, True)
            Else
                Call WriteConsoleMsg(LadrOnIndex, "No has logrado robar nada!", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(VictimaIndex, "" & .Name & " ha intentado robarte!", FontTypeNames.FONTTYPE_INFO)
                
                Call SubirSkill(LadrOnIndex, eSkill.Robar, False)

            End If
        
            If Not criminal(LadrOnIndex) Then
                If Not criminal(VictimaIndex) Then
                    Call VolverCriminal(LadrOnIndex)

                End If

            End If
            
            ' Se pudo haber convertido si robo a un ciuda
            If criminal(LadrOnIndex) Then
                .Reputacion.LadronesRep = .Reputacion.LadronesRep + vlLadron

                If .Reputacion.LadronesRep > MAXREP Then .Reputacion.LadronesRep = MAXREP

            End If

        End If

    End With

    Exit Sub

errHandler:
    Call LogError("Error en DoRobar. Error " & Err.Number & " : " & Err.description)

End Sub

''
' Check if one item is stealable
'
' @param VictimaIndex Specifies reference to victim
' @param Slot Specifies reference to victim's inventory slot
' @return If the item is stealable
Public Function ObjEsRobable(ByVal VictimaIndex As Integer, _
                             ByVal Slot As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    ' Agregue los barcos
    ' Esta funcion determina que objetos son robables.
    ' 22/05/2010: Los items newbies ya no son robables.
    '***************************************************

    Dim OI As Integer

    OI = UserList(VictimaIndex).Invent.Object(Slot).ObjIndex

    ObjEsRobable = ObjData(OI).OBJType <> eOBJType.otLlaves And UserList(VictimaIndex).Invent.Object(Slot).Equipped = 0 And ObjData(OI).Real = 0 And ObjData(OI).Caos = 0 And ObjData(OI).OBJType <> eOBJType.otBarcos And ObjData(OI).OBJType <> eOBJType.otMonturas And ObjData(OI).NoRobable = 1 And Not ItemNewbie(OI)

End Function

''
' Try to steal an item to another character
'
' @param LadrOnIndex Specifies reference to user that stoles
' @param VictimaIndex Specifies reference to user that is being stolen
Public Sub RobarObjeto(ByVal LadrOnIndex As Integer, ByVal VictimaIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: 02/04/2010
    '02/04/2010: ZaMa - Modifico la cantidad de items robables por el ladron.
    '***************************************************

    Dim flag As Boolean

    Dim i    As Integer

    flag = False

    With UserList(VictimaIndex)

        If RandomNumber(1, 12) < 6 Then 'Comenzamos por el principio o el final?
            i = 1

            Do While Not flag And i <= .CurrentInventorySlots

                'Hay objeto en este slot?
                If .Invent.Object(i).ObjIndex > 0 Then
                    If ObjEsRobable(VictimaIndex, i) Then
                        If RandomNumber(1, 10) < 4 Then flag = True

                    End If

                End If

                If Not flag Then i = i + 1
            Loop
        Else
            i = .CurrentInventorySlots

            Do While Not flag And i > 0

                'Hay objeto en este slot?
                If .Invent.Object(i).ObjIndex > 0 Then
                    If ObjEsRobable(VictimaIndex, i) Then
                        If RandomNumber(1, 10) < 4 Then flag = True

                    End If

                End If

                If Not flag Then i = i - 1
            Loop

        End If
    
        If flag Then

            Dim MiObj     As obj

            Dim Num       As Integer

            Dim ObjAmount As Integer
        
            ObjAmount = .Invent.Object(i).Amount
        
            'Cantidad al azar entre el 5% y el 10% del total, con minimo 1.
            Num = MaximoInt(1, RandomNumber(ObjAmount * 0.05, ObjAmount * 0.1))
                                    
            MiObj.Amount = Num
            MiObj.ObjIndex = .Invent.Object(i).ObjIndex
        
            .Invent.Object(i).Amount = ObjAmount - Num
                    
            If .Invent.Object(i).Amount <= 0 Then
                Call QuitarUserInvItem(VictimaIndex, CByte(i), 1)

            End If
                
            Call UpdateUserInv(False, VictimaIndex, CByte(i))
                    
            If Not MeterItemEnInventario(LadrOnIndex, MiObj) Then
                Call TirarItemAlPiso(UserList(LadrOnIndex).Pos, MiObj)

            End If
        
            If UserList(LadrOnIndex).clase = eClass.Thief Then
                Call WriteConsoleMsg(LadrOnIndex, "Has robado " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).Name, FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(LadrOnIndex, "Has hurtado " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).Name, FontTypeNames.FONTTYPE_INFO)

            End If

        Else
            Call WriteConsoleMsg(LadrOnIndex, "No has logrado robar ningun objeto.", FontTypeNames.FONTTYPE_INFO)

        End If

        'If exiting, cancel de quien es robado
        Call CancelExit(VictimaIndex)
        
        'Si esta casteando, lo cancelamos
        Call CancelCast(VictimaIndex)

    End With

End Sub

Public Sub DoApunalar(ByVal userIndex As Integer, _
                      ByVal VictimNpcIndex As Integer, _
                      ByVal VictimUserIndex As Integer, _
                      ByVal dano As Long)

    '***************************************************
    'Autor: Nacho (Integer) & Unknown (orginal version)
    'Last Modification: 04/17/08 - (NicoNZ)
    'Simplifique la cuenta que hacia para sacar la suerte
    'y arregle la cuenta que hacia para sacar el dano
    '***************************************************
    Dim Suerte As Integer

    Dim Skill  As Integer

    Skill = UserList(userIndex).Stats.UserSkills(eSkill.Apunalar)

    Select Case UserList(userIndex).clase

        Case eClass.Assasin
            Suerte = Int(((0.00004 * Skill - 0.002) * Skill + 0.098) * Skill + 4.25)
    
        Case eClass.Cleric, eClass.Paladin, eClass.Pirat
            Suerte = Int(((0.000003 * Skill + 0.0006) * Skill + 0.0107) * Skill + 4.93)
    
        Case eClass.Bard
            Suerte = Int(((0.000002 * Skill + 0.0002) * Skill + 0.032) * Skill + 4.81)
    
        Case Else
            Suerte = Int(0.0361 * Skill + 4.39)

    End Select

    If RandomNumber(0, 100) < Suerte Then
        If VictimUserIndex <> 0 Then
            If UserList(userIndex).clase = eClass.Assasin Then
                dano = Round(dano * 1.4, 0)
            Else
                dano = Round(dano * 1.5, 0)

            End If
        
            With UserList(VictimUserIndex)
                .Stats.MinHp = .Stats.MinHp - dano
                
                'Renderizo el dano en render
                Call SendData(SendTarget.ToPCArea, VictimUserIndex, PrepareMessageCreateDamage(UserList(VictimUserIndex).Pos.X, UserList(VictimUserIndex).Pos.Y, dano, DAMAGE_PUNAL))
                
                Call WriteConsoleMsg(userIndex, "Has apunalado a " & .Name & " por " & dano, FontTypeNames.FONTTYPE_FIGHT)
                Call WriteConsoleMsg(VictimUserIndex, "Te ha apunalado " & UserList(userIndex).Name & " por " & dano, FontTypeNames.FONTTYPE_FIGHT)

            End With
        
        Else
            
            With Npclist(VictimNpcIndex)
                'Si el NPC es un Dummy no aplicamos el daño
                If Not .NPCtype = eNPCType.dummy Then
                    .Stats.MinHp = .Stats.MinHp - Int(dano * 2)
                End If
                
                'Renderizo el dano en render
                Call SendData(SendTarget.ToPCArea, VictimNpcIndex, PrepareMessageCreateDamage(.Pos.X, .Pos.Y, Int(dano * 2), DAMAGE_PUNAL))
                
                Call WriteConsoleMsg(userIndex, "Has apunalado la criatura por " & Int(dano * 2), FontTypeNames.FONTTYPE_FIGHT)
                Call CalcularDarExp(userIndex, VictimNpcIndex, dano * 2)
            
            End With

        End If
    
        Call SubirSkill(userIndex, eSkill.Apunalar, True)
    Else
        Call WriteConsoleMsg(userIndex, "No has logrado apunalar a tu enemigo!", FontTypeNames.FONTTYPE_FIGHT)
        Call SubirSkill(userIndex, eSkill.Apunalar, False)

    End If

End Sub

Public Sub DoAcuchillar(ByVal userIndex As Integer, _
                        ByVal VictimNpcIndex As Integer, _
                        ByVal VictimUserIndex As Integer, _
                        ByVal dano As Integer)
    '***************************************************
    'Autor: ZaMa
    'Last Modification: 12/01/2010
    '***************************************************

    If RandomNumber(1, 100) <= PROB_ACUCHILLAR Then
        dano = Int(dano * DANO_ACUCHILLAR)
        
        If VictimUserIndex <> 0 Then
        
            With UserList(VictimUserIndex)
                .Stats.MinHp = .Stats.MinHp - dano
                Call WriteConsoleMsg(userIndex, "Has acuchillado a " & .Name & " por " & dano, FontTypeNames.FONTTYPE_FIGHT)
                Call WriteConsoleMsg(VictimUserIndex, UserList(userIndex).Name & " te ha acuchillado por " & dano, FontTypeNames.FONTTYPE_FIGHT)

            End With
            
        Else
            With Npclist(VictimNpcIndex)
            
                'Si el NPC es un Dummy no aplicamos el daño
                If Not .NPCtype = eNPCType.dummy Then
                    .Stats.MinHp = .Stats.MinHp - dano
                End If
                
                Call WriteConsoleMsg(userIndex, "Has acuchillado a la criatura por " & dano, FontTypeNames.FONTTYPE_FIGHT)
                Call CalcularDarExp(userIndex, VictimNpcIndex, dano)
            End With
        End If

    End If
    
End Sub

Public Sub DoGolpeCritico(ByVal userIndex As Integer, _
                          ByVal VictimNpcIndex As Integer, _
                          ByVal VictimUserIndex As Integer, _
                          ByVal dano As Long)

    '***************************************************
    'Autor: Pablo (ToxicWaste)
    'Last Modification: 28/01/2007
    '01/06/2010: ZaMa - Valido si tiene arma equipada antes de preguntar si es vikinga.
    '***************************************************
    Dim Suerte      As Integer

    Dim Skill       As Integer

    Dim WeaponIndex As Integer
    
    With UserList(userIndex)

        ' Es bandido?
        If .clase <> eClass.Bandit Then Exit Sub
        
        WeaponIndex = .Invent.WeaponEqpObjIndex
        
        ' Es una espada vikinga?
        If WeaponIndex <> ESPADA_VIKINGA Then Exit Sub
    
        Skill = .Stats.UserSkills(eSkill.Wrestling)

    End With
    
    Suerte = Int((((0.00000003 * Skill + 0.000006) * Skill + 0.000107) * Skill + 0.0893) * 100)
    
    If RandomNumber(1, 100) <= Suerte Then
    
        dano = Int(dano * 0.75)
        
        If VictimUserIndex <> 0 Then
            
            With UserList(VictimUserIndex)
                .Stats.MinHp = .Stats.MinHp - dano
                
                'Renderizo el dano en render
                Call SendData(SendTarget.ToPCArea, VictimUserIndex, PrepareMessageCreateDamage(.Pos.X, .Pos.Y, Int(dano * 2), DAMAGE_PUNAL))
                
                Call WriteConsoleMsg(userIndex, "Has golpeado criticamente a " & .Name & " por " & dano & ".", FontTypeNames.FONTTYPE_FIGHT)
                Call WriteConsoleMsg(VictimUserIndex, UserList(userIndex).Name & " te ha golpeado criticamente por " & dano & ".", FontTypeNames.FONTTYPE_FIGHT)

            End With
            
        Else
            
            With Npclist(VictimNpcIndex)
                'Si el NPC es un Dummy no aplicamos el daño
                If Not .NPCtype = eNPCType.dummy Then
                    .Stats.MinHp = .Stats.MinHp - dano
                End If
                
                'Renderizo el dano en render
                Call SendData(SendTarget.ToPCArea, VictimNpcIndex, PrepareMessageCreateDamage(.Pos.X, .Pos.Y, Int(dano * 2), DAMAGE_PUNAL))
                
                Call WriteConsoleMsg(userIndex, "Has golpeado criticamente a la criatura por " & dano & ".", FontTypeNames.FONTTYPE_FIGHT)
                
                Call CalcularDarExp(userIndex, VictimNpcIndex, dano)
            End With
            
           
            
        End If
        
    End If

End Sub

Public Sub QuitarSta(ByVal userIndex As Integer, ByVal Cantidad As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo errHandler

    UserList(userIndex).Stats.MinSta = UserList(userIndex).Stats.MinSta - Cantidad

    If UserList(userIndex).Stats.MinSta < 0 Then UserList(userIndex).Stats.MinSta = 0
    Call WriteUpdateSta(userIndex)
    
    Exit Sub

errHandler:
    Call LogError("Error en QuitarSta. Error " & Err.Number & " : " & Err.description)
    
End Sub

Public Sub DoTalar(ByVal userIndex As Integer, _
                   Optional ByVal DarMaderaElfica As Boolean = False)

    '***************************************************
    'Autor: Unknown
    'Last Modification: 26/10/2018
    '16/11/2009: ZaMa - Ahora Se puede dar madera elfica.
    '16/11/2009: ZaMa - Implementado nuevo sistema de extraccion.
    '11/05/2010: ZaMa - Arreglo formula de maximo de items contruibles/extraibles.
    '05/13/2010: Pato - Refix a la formula de maximo de items construibles/extraibles.
    '22/05/2010: ZaMa - Los caos ya no suben plebe al trabajar.
    '28/05/2010: ZaMa - Los pks no suben plebe al trabajar.
    '26/10/2018: CHOTS - Multiplicador de oficios
    '***************************************************
    On Error GoTo errHandler

    Dim Suerte        As Integer

    Dim res           As Integer

    Dim MAXITEMS      As Integer

    Dim CantidadItems As Integer

    Dim Skill         As Integer

    With UserList(userIndex)

        Call QuitarSta(userIndex, EsfuerzoExtraer)
    
        Skill = .Stats.UserSkills(eSkill.Talar)
        Suerte = Int(-0.00125 * Skill * Skill - 0.3 * Skill + 49)
    
        res = RandomNumber(1, Suerte)
    
        If res <= DificultadExtraer Then

            Dim MiObj As obj
        
            MAXITEMS = MaxItemsExtraibles(.Stats.ELV)
        
            CantidadItems = RandomNumber(1, MAXITEMS)

            CantidadItems = CantidadItems * OficioMultiplier
            
            With MiObj
                .Amount = CantidadItems
                .ObjIndex = IIf(DarMaderaElfica, LenaElfica, Lena)
            End With
            
        
            If Not MeterItemEnInventario(userIndex, MiObj) Then
                Call TirarItemAlPiso(.Pos, MiObj)

            End If
        
            Call WriteConsoleMsg(userIndex, "Has conseguido algo de lena!", FontTypeNames.FONTTYPE_INFO)
            
            'Renderizo el dano en render.
            Call WriteMessageCreateDamage(userIndex, MiObj.Amount, DAMAGE_TRABAJO)
            
            Call SubirSkill(userIndex, eSkill.Talar, True)
        Else

            '[CDT 17-02-2004]
            If Not .flags.UltimoMensaje = 8 Then
                Call WriteConsoleMsg(userIndex, "No has obtenido lena!", FontTypeNames.FONTTYPE_INFO)
                .flags.UltimoMensaje = 8

            End If

            '[/CDT]
            Call SubirSkill(userIndex, eSkill.Talar, False)

        End If
    
        If Not criminal(userIndex) Then
            .Reputacion.PlebeRep = .Reputacion.PlebeRep + vlProleta

            If .Reputacion.PlebeRep > MAXREP Then .Reputacion.PlebeRep = MAXREP

        End If
    
        .Counters.Trabajando = .Counters.Trabajando + 1

    End With

    Exit Sub

errHandler:
    Call LogError("Error en DoTalar")

End Sub

Public Sub DoMineria(ByVal userIndex As Integer)

    '***************************************************
    'Autor: Unknown
    'Last Modification: 26/10/2018
    '16/11/2009: ZaMa - Implementado nuevo sistema de extraccion.
    '11/05/2010: ZaMa - Arreglo formula de maximo de items contruibles/extraibles.
    '05/13/2010: Pato - Refix a la formula de maximo de items construibles/extraibles.
    '22/05/2010: ZaMa - Los caos ya no suben plebe al trabajar.
    '28/05/2010: ZaMa - Los pks no suben plebe al trabajar.
    '26/10/2018: CHOTS - Multiplicador de oficios
    '***************************************************
    On Error GoTo errHandler

    Dim Suerte        As Integer

    Dim res           As Integer

    Dim MAXITEMS      As Integer

    Dim CantidadItems As Integer

    With UserList(userIndex)

        Call QuitarSta(userIndex, EsfuerzoExtraer)

        Dim Skill As Integer

        Skill = .Stats.UserSkills(eSkill.Mineria)
        Suerte = Int(-0.00125 * Skill * Skill - 0.3 * Skill + 49)
    
        res = RandomNumber(1, Suerte)

        If res <= DificultadExtraer Then

            Dim MiObj As obj
        
            If .flags.TargetObj = 0 Then Exit Sub
        
            MiObj.ObjIndex = ObjData(.flags.TargetObj).RecursoIndex
        
            MAXITEMS = MaxItemsExtraibles(.Stats.ELV)
            
            CantidadItems = RandomNumber(1, MAXITEMS)

            CantidadItems = CantidadItems * OficioMultiplier

            MiObj.Amount = CantidadItems
       
            If Not MeterItemEnInventario(userIndex, MiObj) Then Call TirarItemAlPiso(.Pos, MiObj)
        
            Call WriteConsoleMsg(userIndex, "Has extraido algunos minerales!", FontTypeNames.FONTTYPE_INFO)
            
            'Renderizo el dano en render.
            Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessageCreateDamage(.Pos.X, .Pos.Y, MiObj.Amount, DAMAGE_TRABAJO))
            Call WriteMessageCreateDamage(userIndex, MiObj.Amount, DAMAGE_TRABAJO)
            
            Call SubirSkill(userIndex, eSkill.Mineria, True)
        Else

            '[CDT 17-02-2004]
            If Not .flags.UltimoMensaje = 9 Then
                Call WriteConsoleMsg(userIndex, "No has conseguido nada!", FontTypeNames.FONTTYPE_INFO)
                .flags.UltimoMensaje = 9

            End If

            '[/CDT]
            Call SubirSkill(userIndex, eSkill.Mineria, False)

        End If
    
        If Not criminal(userIndex) Then
            .Reputacion.PlebeRep = .Reputacion.PlebeRep + vlProleta

            If .Reputacion.PlebeRep > MAXREP Then .Reputacion.PlebeRep = MAXREP

        End If
    
        .Counters.Trabajando = .Counters.Trabajando + 1
        
        'Play sound!
        Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessagePlayWave(SND_MINERO, .Pos.X, .Pos.Y))

    End With

    Exit Sub

errHandler:
    Call LogError("Error en Sub DoMineria")

End Sub

Public Sub DoMeditar(ByVal userIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    With UserList(userIndex)
        .Counters.IdleCount = 0
        
        Dim Suerte       As Integer

        Dim res          As Integer

        Dim cant         As Integer

        Dim MeditarSkill As Byte
    
        'Barrin 3/10/03
        'Esperamos a que se termine de concentrar
        Dim TActual      As Long

        TActual = GetTickCount() And &H7FFFFFFF

        If TActual - .Counters.tInicioMeditar < TIEMPO_INICIOMEDITAR Then
            Exit Sub

        End If
        
        If .Counters.bPuedeMeditar = False Then
            .Counters.bPuedeMeditar = True

        End If
            
        If .Stats.MinMAN >= .Stats.MaxMAN Then
            Call WriteConsoleMsg(userIndex, "Has terminado de meditar.", FontTypeNames.FONTTYPE_INFO)
            Call WriteMeditateToggle(userIndex)
            .flags.Meditando = False
            .Char.FX = 0
            .Char.loops = 0
            Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessageCreateFX(.Char.CharIndex, 0, 0))
            Exit Sub

        End If
        
        MeditarSkill = .Stats.UserSkills(eSkill.Meditar)
        
        If MeditarSkill <= 10 Then
            Suerte = 35
        ElseIf MeditarSkill <= 20 Then
            Suerte = 30
        ElseIf MeditarSkill <= 30 Then
            Suerte = 28
        ElseIf MeditarSkill <= 40 Then
            Suerte = 24
        ElseIf MeditarSkill <= 50 Then
            Suerte = 22
        ElseIf MeditarSkill <= 60 Then
            Suerte = 20
        ElseIf MeditarSkill <= 70 Then
            Suerte = 18
        ElseIf MeditarSkill <= 80 Then
            Suerte = 15
        ElseIf MeditarSkill <= 90 Then
            Suerte = 10
        ElseIf MeditarSkill < 100 Then
            Suerte = 7
        Else
            Suerte = 5

        End If

        res = RandomNumber(1, Suerte)
        
        If res = 1 Then
            
            cant = Porcentaje(.Stats.MaxMAN, PorcentajeRecuperoMana)

            If cant <= 0 Then cant = 1
            .Stats.MinMAN = .Stats.MinMAN + cant

            If .Stats.MinMAN > .Stats.MaxMAN Then .Stats.MinMAN = .Stats.MaxMAN
            
            'Renderizo el dano en render.
            Call WriteMessageCreateDamage(userIndex, cant, DAMAGE_TRABAJO)
            
            Call WriteUpdateMana(userIndex)
            Call SubirSkill(userIndex, eSkill.Meditar, True)
        Else
            Call SubirSkill(userIndex, eSkill.Meditar, False)

        End If

    End With

End Sub

Public Sub DoDesequipar(ByVal userIndex As Integer, ByVal victimIndex As Integer)
    '***************************************************
    'Author: ZaMa
    'Last Modif: 15/04/2010
    'Unequips either shield, weapon or helmet from target user.
    '***************************************************

    Dim Probabilidad   As Integer

    Dim Resultado      As Integer

    Dim WrestlingSkill As Byte

    Dim AlgoEquipado   As Boolean
    
    With UserList(userIndex)

        ' Si no tiene guantes de hurto no desequipa.
        If .Invent.AnilloEqpObjIndex <> GUANTE_HURTO Then Exit Sub
        
        ' Si no esta solo con manos, no desequipa tampoco.
        If .Invent.WeaponEqpObjIndex > 0 Then Exit Sub
        
        WrestlingSkill = .Stats.UserSkills(eSkill.Wrestling)
        
        Probabilidad = WrestlingSkill * 0.2 + .Stats.ELV * 0.66

    End With
   
    With UserList(victimIndex)

        ' Si tiene escudo, intenta desequiparlo
        If .Invent.EscudoEqpObjIndex > 0 Then
            
            Resultado = RandomNumber(1, 100)
            
            If Resultado <= Probabilidad Then
                ' Se lo desequipo
                Call Desequipar(victimIndex, .Invent.EscudoEqpSlot)
                
                Call WriteConsoleMsg(userIndex, "Has logrado desequipar el escudo de tu oponente!", FontTypeNames.FONTTYPE_FIGHT)
                
                If .Stats.ELV < 20 Then
                    Call WriteConsoleMsg(victimIndex, "Tu oponente te ha desequipado el escudo!", FontTypeNames.FONTTYPE_FIGHT)

                End If
                
                Exit Sub

            End If
            
            AlgoEquipado = True

        End If
        
        ' No tiene escudo, o fallo desequiparlo, entonces trata de desequipar arma
        If .Invent.WeaponEqpObjIndex > 0 Then
            
            Resultado = RandomNumber(1, 100)
            
            If Resultado <= Probabilidad Then
                ' Se lo desequipo
                Call Desequipar(victimIndex, .Invent.WeaponEqpSlot)
                
                Call WriteConsoleMsg(userIndex, "Has logrado desarmar a tu oponente!", FontTypeNames.FONTTYPE_FIGHT)
                
                If .Stats.ELV < 20 Then
                    Call WriteConsoleMsg(victimIndex, "Tu oponente te ha desarmado!", FontTypeNames.FONTTYPE_FIGHT)

                End If
                
                Exit Sub

            End If
            
            AlgoEquipado = True

        End If
        
        ' No tiene arma, o fallo desequiparla, entonces trata de desequipar casco
        If .Invent.CascoEqpObjIndex > 0 Then
            
            Resultado = RandomNumber(1, 100)
            
            If Resultado <= Probabilidad Then
                ' Se lo desequipo
                Call Desequipar(victimIndex, .Invent.CascoEqpSlot)
                
                Call WriteConsoleMsg(userIndex, "Has logrado desequipar el casco de tu oponente!", FontTypeNames.FONTTYPE_FIGHT)
                
                If .Stats.ELV < 20 Then
                    Call WriteConsoleMsg(victimIndex, "Tu oponente te ha desequipado el casco!", FontTypeNames.FONTTYPE_FIGHT)

                End If
                
                Exit Sub

            End If
            
            AlgoEquipado = True

        End If
    
        If AlgoEquipado Then
            Call WriteConsoleMsg(userIndex, "Tu oponente no tiene equipado items!", FontTypeNames.FONTTYPE_FIGHT)
        Else
            Call WriteConsoleMsg(userIndex, "No has logrado desequipar ningun item a tu oponente!", FontTypeNames.FONTTYPE_FIGHT)

        End If
    
    End With

End Sub

Public Sub DoHurtar(ByVal userIndex As Integer, ByVal VictimaIndex As Integer)

    '***************************************************
    'Author: Pablo (ToxicWaste)
    'Last Modif: 03/03/2010
    'Implements the pick pocket skill of the Bandit :)
    '03/03/2010 - Pato: Solo se puede hurtar si no esta en trigger 6 :)
    '***************************************************
    Dim OtroUserIndex As Integer

    If TriggerZonaPelea(userIndex, VictimaIndex) <> TRIGGER6_AUSENTE Then Exit Sub

    If UserList(userIndex).clase <> eClass.Bandit Then Exit Sub

    'Esto es precario y feo, pero por ahora no se me ocurrio nada mejor.
    'Uso el slot de los anillos para "equipar" los guantes.
    'Y los reconozco porque les puse DefensaMagicaMin y Max = 0
    If UserList(userIndex).Invent.AnilloEqpObjIndex <> GUANTE_HURTO Then Exit Sub

    Dim res As Integer

    res = RandomNumber(1, 100)

    If (res < 20) Then
        If TieneObjetosRobables(VictimaIndex) Then
    
            If UserList(VictimaIndex).flags.Comerciando Then
                OtroUserIndex = UserList(VictimaIndex).ComUsu.DestUsu
                
                If OtroUserIndex > 0 And OtroUserIndex <= MaxUsers Then
                    Call WriteConsoleMsg(VictimaIndex, "Comercio cancelado, te estan robando!!", FontTypeNames.FONTTYPE_WARNING)
                    Call WriteConsoleMsg(OtroUserIndex, "Comercio cancelado por el otro usuario!!", FontTypeNames.FONTTYPE_WARNING)
                
                    Call LimpiarComercioSeguro(VictimaIndex)

                End If

            End If
                
            Call RobarObjeto(userIndex, VictimaIndex)
            Call WriteConsoleMsg(VictimaIndex, "" & UserList(userIndex).Name & " es un Bandido!", FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(userIndex, UserList(VictimaIndex).Name & " no tiene objetos.", FontTypeNames.FONTTYPE_INFO)

        End If

    End If

End Sub

Public Sub DoHandInmo(ByVal userIndex As Integer, ByVal VictimaIndex As Integer)

    '***************************************************
    'Author: Pablo (ToxicWaste)
    'Last Modif: 17/02/2007
    'Implements the special Skill of the Thief
    '***************************************************
    If UserList(VictimaIndex).flags.Paralizado = 1 Then Exit Sub
    If UserList(userIndex).clase <> eClass.Thief Then Exit Sub
    
    If UserList(userIndex).Invent.AnilloEqpObjIndex <> GUANTE_HURTO Then Exit Sub
        
    Dim res As Integer

    res = RandomNumber(0, 100)

    If res < (UserList(userIndex).Stats.UserSkills(eSkill.Wrestling) / 4) Then
        UserList(VictimaIndex).flags.Paralizado = 1
        UserList(VictimaIndex).Counters.Paralisis = IntervaloParalizado / 2
        
        UserList(VictimaIndex).flags.ParalizedByIndex = userIndex
        UserList(VictimaIndex).flags.ParalizedBy = UserList(userIndex).Name
        
        Call WriteParalizeOK(VictimaIndex)
        Call WriteConsoleMsg(userIndex, "Tu golpe ha dejado inmovil a tu oponente", FontTypeNames.FONTTYPE_FIGHT)
        Call WriteConsoleMsg(VictimaIndex, "El golpe te ha dejado inmovil!", FontTypeNames.FONTTYPE_FIGHT)

    End If

End Sub

Public Sub Desarmar(ByVal userIndex As Integer, ByVal victimIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: 02/04/2010 (ZaMa)
    '02/04/2010: ZaMa - Nueva formula para desarmar.
    '***************************************************

    Dim Probabilidad   As Integer

    Dim Resultado      As Integer

    Dim WrestlingSkill As Byte
    
    With UserList(userIndex)
        WrestlingSkill = .Stats.UserSkills(eSkill.Wrestling)
        
        Probabilidad = WrestlingSkill * 0.2 + .Stats.ELV * 0.66
        
        Resultado = RandomNumber(1, 100)
        
        If Resultado <= Probabilidad Then
            Call Desequipar(victimIndex, UserList(victimIndex).Invent.WeaponEqpSlot)
            Call WriteConsoleMsg(userIndex, "Has logrado desarmar a tu oponente!", FontTypeNames.FONTTYPE_FIGHT)

            If UserList(victimIndex).Stats.ELV < 20 Then
                Call WriteConsoleMsg(victimIndex, "Tu oponente te ha desarmado!", FontTypeNames.FONTTYPE_FIGHT)

            End If

        End If

    End With
    
End Sub

Public Function MaxItemsConstruibles(ByVal userIndex As Integer) As Integer
    '***************************************************
    'Author: ZaMa
    'Last Modification: 29/01/2010
    '11/05/2010: ZaMa - Arreglo formula de maximo de items contruibles/extraibles.
    '05/13/2010: Pato - Refix a la formula de maximo de items construibles/extraibles.
    '***************************************************
    
    With UserList(userIndex)

    MaxItemsConstruibles = MaximoInt(1, CInt((.Stats.ELV - 2) * 0.2))

    End With

End Function

Public Function MaxItemsExtraibles(ByVal UserLevel As Integer) As Integer
    '***************************************************
    'Author: ZaMa
    'Last Modification: 14/05/2010
    '***************************************************
    MaxItemsExtraibles = MaximoInt(1, CInt((UserLevel - 2) * 0.2)) + 1

End Function

Public Sub ImitateNpc(ByVal userIndex As Integer, ByVal NPCIndex As Integer)
    '***************************************************
    'Author: ZaMa
    'Last Modification: 20/11/2010
    'Copies body, head and desc from previously clicked npc.
    '***************************************************
    
    With UserList(userIndex)
        
        ' Copy desc
        .DescRM = Npclist(NPCIndex).Name
        
        ' Remove Anims (Npcs don't use equipment anims yet)
        .Char.CascoAnim = NingunCasco
        .Char.ShieldAnim = NingunEscudo
        .Char.WeaponAnim = NingunArma
        
        ' If admin is invisible the store it in old char
        If .flags.AdminInvisible = 1 Or .flags.invisible = 1 Or .flags.Oculto = 1 Then
            
            .flags.OldBody = Npclist(NPCIndex).Char.body
            .flags.OldHead = Npclist(NPCIndex).Char.Head
        Else
            .Char.body = Npclist(NPCIndex).Char.body
            .Char.Head = Npclist(NPCIndex).Char.Head
            
            Call ChangeUserChar(userIndex, .Char.body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraAnim, .Char.AuraColor)

        End If
    
    End With
    
End Sub

Public Sub DoEquita(ByVal userIndex As Integer, _
                    ByRef Montura As ObjData, _
                    ByVal Slot As Integer)
    '***************************************************
    'Author: Recox
    'Last Modification: 06/04/2020
    'Podemos usar monturas ahora
    '06/04/2020: FrankoH298 - Ahora hay un timer para poder montarte
    '***************************************************

    With UserList(userIndex)
    
        If UserList(userIndex).Stats.UserSkills(Equitacion) < Montura.MinSkill Then
            Call WriteConsoleMsg(userIndex, "Para usar esta montura necesitas " & Montura.MinSkill & " puntos en equitación.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        '¿Esta intentando usar una montura de tipo dungeon fuera de un dungeon?
        If MapInfo(.Pos.Map).Zona <> "DUNGEON" And Montura.MontTipo = 1 Then
            Call WriteConsoleMsg(userIndex, "No puedes utilizar esta montura fuera de un dungeon.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        '¿Esta en un dungeon y la montura no es de tipo dungeon?
        If MapInfo(.Pos.Map).Zona = "DUNGEON" And Montura.MontTipo <> 1 Then
            Call WriteConsoleMsg(userIndex, "No puedes utilizar esta montura en dungeon.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(userIndex, "No puedes utilizar la montura mientras estas muerto !!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        If .flags.Navegando = 1 Then
            Call WriteConsoleMsg(userIndex, "No puedes utilizar la montura mientras navegas !!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If

        If MapData(.Pos.Map, .Pos.X, .Pos.Y).Trigger = eTrigger.BAJOTECHO Or MapData(.Pos.Map, .Pos.X, .Pos.Y).Trigger = eTrigger.CASA Then
            'TODO: SACAR ESTA VALIDACION DE ACA, Y HACER UN legalpos HAY TECHO en el cliente
            If .flags.Equitando = 0 Then Exit Sub

            Call WriteConsoleMsg(userIndex, "No puedes utilizar la montura bajo techo!", FontTypeNames.FONTTYPE_INFO)
        End If

        ' If .flags.Metamorfosis = 1 Then 'Metamorfosis
        '     Call WriteConsoleMsg(UserIndex, "No puedes montar mientras estas metamorfoseado.", FontTypeNames.FONTTYPE_INFO)
        '     Exit Sub
        ' End If

        ' No estaba equitando
        If .flags.Equitando = 0 Then

            If .Counters.MonturaCounter <= 0 Then
                .Invent.MonturaObjIndex = .Invent.Object(Slot).ObjIndex
                .Invent.MonturaEqpSlot = Slot
    
                Call ToggleMonturaBody(userIndex)
                Call SetVisibleStateForUserAfterNavigateOrEquitate(userIndex)
    
                '  Comienza a equitar
                .flags.Equitando = 1

                Call WriteEquitandoToggle(userIndex)

                'Mostramos solo el casco de los items equipados por que los demas items quedan mal en el render, solo es un tema visual (Recox)
                Call ChangeUserChar(userIndex, .Char.body, .Char.Head, .Char.Heading, NingunArma, NingunEscudo, .Char.CascoAnim, .Char.AuraAnim, .Char.AuraColor)
            Else
                Call WriteConsoleMsg(userIndex, "Debe esperar " & .Counters.MonturaCounter & " segundos para volver a usar tu montura", FontTypeNames.FONTTYPE_INFO)
            End If
            
        ' Estaba equitando
        Else
            Call UnmountMontura(userIndex)
            Call WriteEquitandoToggle(userIndex)

        End If


    End With

End Sub

Public Sub UnmountMontura(ByVal userIndex As Integer)
    With UserList(userIndex)
        .Invent.MonturaObjIndex = 0
        .Invent.MonturaEqpSlot = 0

        .Char.Head = .OrigChar.Head

        ' Seteamos el equipo que tiene y lo mostramos en el render.
        Call SetEquipmentOnCharAfterNavigateOrEquitate(userIndex)
        Call ChangeUserChar(userIndex, .Char.body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim, .Char.AuraAnim, .Char.AuraColor)
  
        ' Termina de equitar
        .flags.Equitando = 0
        .Counters.MonturaCounter = 3

    End With
End Sub

Private Sub SetVisibleStateForUserAfterNavigateOrEquitate(ByVal userIndex As Integer)

    With UserList(userIndex)

        ' Pierde el ocultar
        If .flags.Oculto = 1 Then
            .flags.Oculto = 0
            .Counters.Ocultando = 0
            Call SetInvisible(userIndex, .Char.CharIndex, False)
            Call WriteConsoleMsg(userIndex, "Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)
        End If

        ' Siempre se ve la montura (Nunca esta invisible), pero solo para el cliente.
        If .flags.invisible = 1 Then
            Call SetInvisible(userIndex, .Char.CharIndex, False)
        End If

    End With

End Sub

Private Sub SetEquipmentOnCharAfterNavigateOrEquitate(ByVal userIndex As Integer)

    With UserList(userIndex)

        If .Invent.ArmourEqpObjIndex > 0 Then
            .Char.body = ObjData(.Invent.ArmourEqpObjIndex).Ropaje
        Else
            Call DarCuerpoDesnudo(userIndex)

        End If
        
        If .Invent.EscudoEqpObjIndex > 0 Then .Char.ShieldAnim = ObjData(.Invent.EscudoEqpObjIndex).ShieldAnim

        If .Invent.WeaponEqpObjIndex > 0 Then .Char.WeaponAnim = GetWeaponAnim(userIndex, .Invent.WeaponEqpObjIndex)

        If .Invent.CascoEqpObjIndex > 0 Then .Char.CascoAnim = ObjData(.Invent.CascoEqpObjIndex).CascoAnim
        
    End With


End Sub

Public Sub DoExtraer(ByVal userIndex As Integer, ByVal Profesion As Integer)

    '***************************************************
    'Autor: Lorwik
    'Fecha: 19/08/2020
    'Descripción: Extrae recursos de forma pasiva
    '***************************************************
    
    On Error GoTo errHandler

    Dim Suerte        As Integer
    Dim res           As Integer
    Dim MAXITEMS      As Integer
    Dim CantidadItems As Integer
    Dim MiObj As obj
    

    With UserList(userIndex)

        If .flags.TargetObj = 0 Then Exit Sub

        '¿La herramienta es de la misma categoria o superior?
        If ObjData(.flags.TargetObj).Recurso.Categoria > ObjData(.Invent.WeaponEqpObjIndex).Herramienta.Categoria Then
            Call WriteConsoleMsg(userIndex, "El recurso que intentas extraer es demasiado duro para esa herramienta.", FontTypeNames.FONTTYPE_INFO)
            Call DejardeTrabajar(userIndex) 'Paramos el macro
            Exit Sub
        End If

        Call QuitarSta(userIndex, EsfuerzoExtraer)

        Dim Skill As Integer

        Skill = .Stats.UserSkills(Profesion)
        Suerte = Int(-0.00125 * Skill * Skill - 0.3 * Skill + 49)
    
        res = RandomNumber(1, Suerte)

        If res <= DificultadExtraer Then
        
            MiObj.ObjIndex = ObjData(.flags.TargetObj).RecursoIndex
        
            MAXITEMS = MaxItemsExtraibles(.Stats.ELV)
            
            CantidadItems = RandomNumber(1, MAXITEMS)

            CantidadItems = CantidadItems * OficioMultiplier

            MiObj.Amount = CantidadItems
       
            If Not MeterItemEnInventario(userIndex, MiObj) Then Call TirarItemAlPiso(.Pos, MiObj)
        
            Call WriteConsoleMsg(userIndex, "Has extraido algunos materiales!", FontTypeNames.FONTTYPE_INFO)
            
            'Renderizo el dano en render.
            Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessageCreateDamage(.Pos.X, .Pos.Y, MiObj.Amount, DAMAGE_TRABAJO))
            Call WriteMessageCreateDamage(userIndex, MiObj.Amount, DAMAGE_TRABAJO)
            
            Call SubirSkill(userIndex, Profesion, True)
        Else

            '[CDT 17-02-2004]
            If Not .flags.UltimoMensaje = 9 Then
                Call WriteConsoleMsg(userIndex, "No has conseguido nada!", FontTypeNames.FONTTYPE_INFO)
                .flags.UltimoMensaje = 9

            End If

            '[/CDT]
            Call SubirSkill(userIndex, Profesion, False)

        End If
    
        If Not criminal(userIndex) Then
            .Reputacion.PlebeRep = .Reputacion.PlebeRep + vlProleta

            If .Reputacion.PlebeRep > MAXREP Then .Reputacion.PlebeRep = MAXREP

        End If
    
        .Counters.Trabajando = .Counters.Trabajando + 1
        
        'Play sound!
        If Profesion = eSkill.Mineria Then
            Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessagePlayWave(SND_MINERO, .Pos.X, .Pos.Y))
            
        ElseIf Profesion = eSkill.Talar Then
            Call SendData(SendTarget.ToPCArea, userIndex, PrepareMessagePlayWave(SND_TALAR, .Pos.X, .Pos.Y))
            
        End If

    End With

    Exit Sub

errHandler:
    Call LogError("Error en Sub DoExtraer")

End Sub
