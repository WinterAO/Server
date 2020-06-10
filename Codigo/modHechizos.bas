Attribute VB_Name = "modHechizos"
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

Sub NpcLanzaSpellSobreUser(ByVal NpcIndex As Integer, _
                           ByVal UserIndex As Integer, _
                           ByVal spell As Integer, _
                           Optional ByVal DecirPalabras As Boolean = False, _
                           Optional ByVal IgnoreVisibilityCheck As Boolean = False)
    '***************************************************
    'Autor: Unknown (orginal version)
    'Last Modification: 06/04/2020
    '13/02/2009: ZaMa - Los npcs que tiren magias, no podran hacerlo en mapas donde no se permita usarla.
    '13/07/2010: ZaMa - Ahora no se contabiliza la muerte de un atacable.
    '21/09/2010: ZaMa - Amplio los tipos de hechizos que pueden lanzar los npcs.
    '21/09/2010: ZaMa - Permito que se ignore el chequeo de visibilidad (pueden atacar a invis u ocultos).
    '11/11/2010: ZaMa - No se envian los efectos del hechizo si no lo castea.
    '06/04/2020: FrankoH298 - Si te lanzan un hechizo te desmonta
    '***************************************************

    If Not IntervaloPermiteAtacarNpc(NpcIndex) Then Exit Sub

    With UserList(UserIndex)
    
        '<<<< Equitando >>>
        If .flags.Equitando = 1 Then
            Call UnmountMontura(UserIndex)
            Call WriteEquitandoToggle(UserIndex)
            
        End If
        
        ' Doesn't consider if the user is hidden/invisible or not.
        If Not IgnoreVisibilityCheck Then
            If UserList(UserIndex).flags.invisible = 1 Or UserList(UserIndex).flags.Oculto = 1 Then Exit Sub

        End If
        
        ' Si no se peude usar magia en el mapa, no le deja hacerlo.
        If MapInfo(UserList(UserIndex).Pos.Map).MagiaSinEfecto > 0 Then Exit Sub

        Dim dano As Integer
    
        ' Heal HP
        If Hechizos(spell).SubeHP = 1 Then
        
            Call SendSpellEffects(UserIndex, NpcIndex, spell, DecirPalabras)
        
            dano = RandomNumber(Hechizos(spell).MinHp, Hechizos(spell).MaxHp)
        
            .Stats.MinHp = .Stats.MinHp + dano

            If .Stats.MinHp > .Stats.MaxHp Then .Stats.MinHp = .Stats.MaxHp
            
            Call WriteConsoleMsg(UserIndex, Npclist(NpcIndex).Name & " te ha quitado " & dano & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
            
            Call WriteUpdateUserStats(UserIndex)
        
            ' Damage
        ElseIf Hechizos(spell).SubeHP = 2 Then
            
            If .flags.Privilegios And PlayerType.User Then
            
                Call SendSpellEffects(UserIndex, NpcIndex, spell, DecirPalabras)
            
                dano = RandomNumber(Hechizos(spell).MinHp, Hechizos(spell).MaxHp)
                
                If .Invent.CascoEqpObjIndex > 0 Then
                    dano = dano - RandomNumber(ObjData(.Invent.CascoEqpObjIndex).DefensaMagicaMin, ObjData(.Invent.CascoEqpObjIndex).DefensaMagicaMax)

                End If
                
                If .Invent.AnilloEqpObjIndex > 0 Then
                    dano = dano - RandomNumber(ObjData(.Invent.AnilloEqpObjIndex).DefensaMagicaMin, ObjData(.Invent.AnilloEqpObjIndex).DefensaMagicaMax)

                End If
                
                If dano < 0 Then dano = 0
            
                .Stats.MinHp = .Stats.MinHp - dano
                
                Call WriteConsoleMsg(UserIndex, Npclist(NpcIndex).Name & " te ha quitado " & dano & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
                
                'Renderizo el dano en render.
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateDamage(.Pos.X, .Pos.Y, dano, DAMAGE_NORMAL))
                
                Call WriteUpdateUserStats(UserIndex)
                
                'Muere
                If .Stats.MinHp < 1 Then
                    .Stats.MinHp = 0

                    If Npclist(NpcIndex).NPCtype = eNPCType.GuardiaReal Then
                        RestarCriminalidad (UserIndex)

                    End If
                    
                    Dim MasterIndex As Integer

                    MasterIndex = Npclist(NpcIndex).MaestroUser
                    
                    '[Barrin 1-12-03]
                    If MasterIndex > 0 Then
                        
                        ' No son frags los muertos atacables
                        If .flags.AtacablePor <> MasterIndex Then
                            'Store it!
                            Call Statistics.StoreFrag(MasterIndex, UserIndex)
                            
                            Call ContarMuerte(UserIndex, MasterIndex)

                        End If
                        
                        Call ActStats(UserIndex, MasterIndex)

                    End If

                    '[/Barrin]
                    
                    Call UserDie(UserIndex)
                    
                End If
            
            End If
            
        End If
        
        ' Paralisis/Inmobilize
        If Hechizos(spell).Paraliza = 1 Or Hechizos(spell).Inmoviliza = 1 Then
        
            If .flags.Paralizado = 0 Then
                
                Call SendSpellEffects(UserIndex, NpcIndex, spell, DecirPalabras)
                
                If .Invent.AnilloEqpObjIndex > 0 Then
                    If ObjData(.Invent.AnilloEqpObjIndex).ImpideParalizar Then
                        Call WriteConsoleMsg(UserIndex, "Tu anillo rechaza los efectos del hechizo.", FontTypeNames.FONTTYPE_FIGHT)
                        Exit Sub
                    End If
                End If
                
                If Hechizos(spell).Inmoviliza = 1 Then
                    .flags.Inmovilizado = 1

                End If
                  
                .flags.Paralizado = 1
                .Counters.Paralisis = IntervaloParalizado
                  
                Call WriteParalizeOK(UserIndex)
                
            End If
            
        End If
        
        ' Stupidity
        If Hechizos(spell).Estupidez = 1 Then
             
            If .flags.Estupidez = 0 Then
            
                Call SendSpellEffects(UserIndex, NpcIndex, spell, DecirPalabras)
            
                If .Invent.AnilloEqpObjIndex > 0 Then
                    If ObjData(.Invent.AnilloEqpObjIndex).ImpideAturdir Then
                        Call WriteConsoleMsg(UserIndex, "Tu anillo rechaza los efectos del hechizo.", FontTypeNames.FONTTYPE_FIGHT)
                        Exit Sub
                    End If
                End If
                  
                .flags.Estupidez = 1
                .Counters.Ceguera = IntervaloInvisible
                          
                Call WriteDumb(UserIndex)
                
            End If

        End If
        
        ' Blind
        If Hechizos(spell).Ceguera = 1 Then
             
            If .flags.Ceguera = 0 Then
            
                Call SendSpellEffects(UserIndex, NpcIndex, spell, DecirPalabras)
            
                If .Invent.AnilloEqpObjIndex > 0 Then
                    If ObjData(.Invent.AnilloEqpObjIndex).ImpideCegar Then
                        Call WriteConsoleMsg(UserIndex, "Tu anillo rechaza los efectos del hechizo.", FontTypeNames.FONTTYPE_FIGHT)
                        Exit Sub
                    End If
                End If
                  
                .flags.Ceguera = 1
                .Counters.Ceguera = IntervaloInvisible
                          
                Call WriteBlind(UserIndex)
                
            End If

        End If
        
        ' Remove Invisibility/Hidden
        If Hechizos(spell).RemueveInvisibilidadParcial = 1 Then
                 
            Call SendSpellEffects(UserIndex, NpcIndex, spell, DecirPalabras)
                 
            'Sacamos el efecto de ocultarse
            If .flags.Oculto = 1 Then
                .Counters.TiempoOculto = 0
                .flags.Oculto = 0
                Call SetInvisible(UserIndex, .Char.CharIndex, False)
                Call WriteConsoleMsg(UserIndex, "Has sido detectado!", FontTypeNames.FONTTYPE_VENENO)
            Else
                'sino, solo lo "iniciamos" en la sacada de invisibilidad.
                Call WriteConsoleMsg(UserIndex, "Comienzas a hacerte visible.", FontTypeNames.FONTTYPE_VENENO)
                .Counters.Invisibilidad = IntervaloInvisible - 1

            End If
        
        End If
        
    End With
    
End Sub

Private Sub SendSpellEffects(ByVal UserIndex As Integer, _
                             ByVal NpcIndex As Integer, _
                             ByVal spell As Integer, _
                             ByVal DecirPalabras As Boolean)

    '***************************************************
    'Author: ZaMa
    'Last Modification: 28/12/2016
    'Sends spell's wav, fx and mgic words to users.
    ' Shak: Palabras magicas
    '***************************************************
    With UserList(UserIndex)
        ' Spell Wav
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(Hechizos(spell).WAV, .Pos.X, .Pos.Y))
            
        ' Spell FX
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(.Char.CharIndex, Hechizos(spell).FXgrh, Hechizos(spell).loops))
    
        ' Spell Words
        If DecirPalabras Then
            Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePalabrasMagicas(spell, Npclist(NpcIndex).Char.CharIndex))

        End If

    End With

End Sub

Public Sub NpcLanzaSpellSobreNpc(ByVal NpcIndex As Integer, _
                                 ByVal TargetNPC As Integer, _
                                 ByVal spellIndex As Integer, _
                                 Optional ByVal DecirPalabras As Boolean = False)
    '***************************************************
    'Author: Unknown
    'Last Modification: 21/09/2010
    '21/09/2010: ZaMa - Now npcs can cast a wider range of spells.
    '***************************************************

    If Not IntervaloPermiteAtacarNpc(NpcIndex) Then Exit Sub
    
    Dim Danio As Integer
    
    With Npclist(TargetNPC)
    
        ' Spell sound and FX
        Call SendData(SendTarget.ToNPCArea, TargetNPC, PrepareMessagePlayWave(Hechizos(spellIndex).WAV, .Pos.X, .Pos.Y))
            
        Call SendData(SendTarget.ToNPCArea, TargetNPC, PrepareMessageCreateFX(.Char.CharIndex, Hechizos(spellIndex).FXgrh, Hechizos(spellIndex).loops))
    
        ' Decir las palabras magicas?
        If DecirPalabras Then
            Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePalabrasMagicas(spellIndex, Npclist(NpcIndex).Char.CharIndex))

        End If
    
        ' Spell deals damage??
        If Hechizos(spellIndex).SubeHP = 2 Then
            
            Danio = RandomNumber(Hechizos(spellIndex).MinHp, Hechizos(spellIndex).MaxHp)
            
            ' Deal damage
            .Stats.MinHp = .Stats.MinHp - Danio
            
            'Muere?
            If .Stats.MinHp < 1 Then
                .Stats.MinHp = 0

                If Npclist(NpcIndex).MaestroUser > 0 Then
                    Call MuereNpc(TargetNPC, Npclist(NpcIndex).MaestroUser)
                Else
                    Call MuereNpc(TargetNPC, 0)

                End If

            End If
            
            ' Spell recovers health??
        ElseIf Hechizos(spellIndex).SubeHP = 1 Then
            
            Danio = RandomNumber(Hechizos(spellIndex).MinHp, Hechizos(spellIndex).MaxHp)
            
            ' Recovers health
            .Stats.MinHp = .Stats.MinHp + Danio
            
            If .Stats.MinHp > .Stats.MaxHp Then
                .Stats.MinHp = .Stats.MaxHp

            End If
            
        End If
        
        ' Spell Adds/Removes poison?
        If Hechizos(spellIndex).Envenena = 1 Then
            .flags.Envenenado = 1
        ElseIf Hechizos(spellIndex).CuraVeneno = 1 Then
            .flags.Envenenado = 0

        End If

        ' Spells Adds/Removes Paralisis/Inmobility?
        If Hechizos(spellIndex).Paraliza = 1 Then
            .flags.Paralizado = 1
            .flags.Inmovilizado = 0
            .Contadores.Paralisis = IntervaloParalizado
            
        ElseIf Hechizos(spellIndex).Inmoviliza = 1 Then
            .flags.Inmovilizado = 1
            .flags.Paralizado = 0
            .Contadores.Paralisis = IntervaloParalizado
            
        ElseIf Hechizos(spellIndex).RemoverParalisis = 1 Then

            If .flags.Paralizado = 1 Or .flags.Inmovilizado = 1 Then
                .flags.Paralizado = 0
                .flags.Inmovilizado = 0
                .Contadores.Paralisis = 0

            End If

        End If
    
    End With

End Sub

Function TieneHechizo(ByVal i As Integer, ByVal UserIndex As Integer) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo Errhandler
    
    Dim j As Integer

    For j = 1 To MAXUSERHECHIZOS

        If UserList(UserIndex).Stats.UserHechizos(j) = i Then
            TieneHechizo = True
            Exit Function

        End If

    Next

    Exit Function
Errhandler:

End Function

Sub AgregarHechizo(ByVal UserIndex As Integer, ByVal Slot As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim hIndex As Integer

    Dim j      As Integer

    With UserList(UserIndex)
        hIndex = ObjData(.Invent.Object(Slot).ObjIndex).HechizoIndex
    
        If Not TieneHechizo(hIndex, UserIndex) Then

            'Buscamos un slot vacio
            For j = 1 To MAXUSERHECHIZOS

                If .Stats.UserHechizos(j) = 0 Then Exit For
            Next j
            
            If .Stats.UserHechizos(j) <> 0 Then
                Call WriteConsoleMsg(UserIndex, "No tienes espacio para mas hechizos.", FontTypeNames.FONTTYPE_INFO)
            Else
                .Stats.UserHechizos(j) = hIndex
                Call UpdateUserHechizos(False, UserIndex, CByte(j))
                'Quitamos del inv el item
                Call QuitarUserInvItem(UserIndex, CByte(Slot), 1)

            End If

        Else
            Call WriteConsoleMsg(UserIndex, "Ya tienes ese hechizo.", FontTypeNames.FONTTYPE_INFO)

        End If

    End With

End Sub
            
Sub DecirPalabrasMagicas(ByVal spellIndex As Integer, ByVal UserIndex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: 17/11/2009
    '25/07/2009: ZaMa - Invisible admins don't say any word when casting a spell
    '17/11/2009: ZaMa - Now the user become visible when casting a spell, if it is hidden
    '28/12/2016: Shak - Palabras magicas
    '21/02/2019: Jopi - Amuleto del Silencio
    '***************************************************
    On Error GoTo Errhandler
    
    ' Amuleto del Silencio
    If TieneObjetos(AMULETO_DEL_SILENCIO, 1, UserIndex) Then Exit Sub
              
    With UserList(UserIndex)

        If .flags.AdminInvisible <> 1 Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePalabrasMagicas(spellIndex, .Char.CharIndex))
                
            ' Si estaba oculto, se vuelve visible
            If .flags.Oculto = 1 Then
                .flags.Oculto = 0
                .Counters.TiempoOculto = 0
                
                If .flags.invisible = 0 Then
                    Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
                    Call SetInvisible(UserIndex, .Char.CharIndex, False)

                End If

            End If

        End If

    End With
    
    Exit Sub
    
Errhandler:
    Call LogError("Error en DecirPalabrasMagicas. Error: " & Err.Number & " - " & Err.description)

End Sub

''
' Check if an user can cast a certain spell
'
' @param UserIndex Specifies reference to user
' @param HechizoIndex Specifies reference to spell
' @return   True if the user can cast the spell, otherwise returns false
Function PuedeLanzar(ByVal UserIndex As Integer, ByVal HechizoIndex As Integer) As Boolean

    '***************************************************
    'Author: Unknown
    'Last Modification: 12/01/2010
    'Last Modification By: ZaMa
    '06/11/09 - Corregida la bonificacion de mana del mimetismo en el druida con flauta magica equipada.
    '19/11/2009: ZaMa - Validacion de mana para el Invocar Mascotas
    '12/01/2010: ZaMa - Validacion de mana para hechizos lanzados por druida.
    '***************************************************
    Dim DruidManaBonus As Single

    With UserList(UserIndex)

        If .flags.Muerto Then
            'Call WriteConsoleMsg(UserIndex, "No puedes lanzar hechizos estando muerto.", FontTypeNames.FONTTYPE_INFO)
            Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
            Exit Function

        End If
            
        If Hechizos(HechizoIndex).NeedStaff > 0 Then
            If .clase = eClass.Mage Then
                If .Invent.WeaponEqpObjIndex > 0 Then
                    If ObjData(.Invent.WeaponEqpObjIndex).StaffPower < Hechizos(HechizoIndex).NeedStaff Then
                        Call WriteConsoleMsg(UserIndex, "No posees un baculo lo suficientemente poderoso para poder lanzar el conjuro.", FontTypeNames.FONTTYPE_INFO)
                        Exit Function

                    End If

                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes lanzar este conjuro sin la ayuda de un baculo.", FontTypeNames.FONTTYPE_INFO)
                    Exit Function

                End If

            End If

        End If
            
        If .Stats.UserSkills(eSkill.Magia) < Hechizos(HechizoIndex).MinSkill Then
            Call WriteConsoleMsg(UserIndex, "No tienes suficientes puntos de magia para lanzar este hechizo.", FontTypeNames.FONTTYPE_INFO)
            Exit Function

        End If
        
        If .Stats.MinSta < Hechizos(HechizoIndex).StaRequerido Then
            If .Genero = eGenero.Hombre Then
                Call WriteConsoleMsg(UserIndex, "Estas muy cansado para lanzar este hechizo.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(UserIndex, "Estas muy cansada para lanzar este hechizo.", FontTypeNames.FONTTYPE_INFO)

            End If

            Exit Function

        End If
    
        DruidManaBonus = 1

        If .clase = eClass.Druid Then
            If .Invent.AnilloEqpObjIndex = FLAUTAELFICA Then

                ' 50% menos de mana requerido para mimetismo
                If Hechizos(HechizoIndex).Mimetiza = 1 Then
                    DruidManaBonus = 0.5
                    
                    ' 30% menos de mana requerido para invocaciones
                ElseIf Hechizos(HechizoIndex).Tipo = uInvocacion Then
                    DruidManaBonus = 0.7
                
                    ' 10% menos de mana requerido para las demas magias, excepto apoca
                ElseIf HechizoIndex <> APOCALIPSIS_SPELL_INDEX Then
                    DruidManaBonus = 0.9

                End If

            End If
            
            ' Necesita tener la barra de mana completa para invocar una mascota
            If Hechizos(HechizoIndex).Warp = 1 Then
                If .Stats.MinMAN <> .Stats.MaxMAN Then
                    Call WriteConsoleMsg(UserIndex, "Debes poseer toda tu mana para poder lanzar este hechizo.", FontTypeNames.FONTTYPE_INFO)
                    Exit Function
                    ' Si no tiene mascotas, no tiene sentido que lo use
                ElseIf .NroMascotas = 0 Then
                    Call WriteConsoleMsg(UserIndex, "Debes poseer alguna mascota para poder lanzar este hechizo.", FontTypeNames.FONTTYPE_INFO)
                    Exit Function

                End If

            End If

        End If
        
        If .Stats.MinMAN < Hechizos(HechizoIndex).ManaRequerido * DruidManaBonus Then
            Call WriteConsoleMsg(UserIndex, "No tienes suficiente mana.", FontTypeNames.FONTTYPE_INFO)
            Exit Function

        End If
        
    End With
    
    PuedeLanzar = True

End Function

Sub HechizoTerrenoEstado(ByVal UserIndex As Integer, ByRef b As Boolean)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim PosCasteadaX As Integer

    Dim PosCasteadaY As Integer

    Dim PosCasteadaM As Integer

    Dim h            As Integer

    Dim tempX        As Integer

    Dim tempY        As Integer

    With UserList(UserIndex)
        PosCasteadaX = .flags.TargetX
        PosCasteadaY = .flags.TargetY
        PosCasteadaM = .flags.TargetMap
        
        h = .flags.Hechizo
        
        If Hechizos(h).RemueveInvisibilidadParcial = 1 Then
            b = True

            For tempX = PosCasteadaX - 8 To PosCasteadaX + 8
                For tempY = PosCasteadaY - 8 To PosCasteadaY + 8

                    If InMapBounds(PosCasteadaM, tempX, tempY) Then
                        If MapData(PosCasteadaM, tempX, tempY).UserIndex > 0 Then

                            'hay un user
                            If UserList(MapData(PosCasteadaM, tempX, tempY).UserIndex).flags.invisible = 1 And UserList(MapData(PosCasteadaM, tempX, tempY).UserIndex).flags.AdminInvisible = 0 Then
                                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(MapData(PosCasteadaM, tempX, tempY).UserIndex).Char.CharIndex, Hechizos(h).FXgrh, Hechizos(h).loops))

                            End If

                        End If

                    End If

                Next tempY
            Next tempX
        
            Call InfoHechizo(UserIndex)

        End If

    End With

End Sub

''
' Le da propiedades al nuevo npc
'
' @param UserIndex  Indice del usuario que invoca.
' @param b  Indica si se termino la operacion.

Sub HechizoInvocacion(ByVal UserIndex As Integer, ByRef HechizoCasteado As Boolean)
    '***************************************************
    'Author: Uknown
    'Last modification: 18/09/2010
    'Sale del sub si no hay una posicion valida.
    '18/11/2009: Optimizacion de codigo.
    '18/09/2010: ZaMa - No se permite invocar en mapas con InvocarSinEfecto.
    '***************************************************

    On Error GoTo Error

    With UserList(UserIndex)

        Dim Mapa As Integer

        Mapa = .Pos.Map
    
        'No permitimos se invoquen criaturas en zonas seguras
        If MapInfo(Mapa).Pk = False Or MapData(Mapa, .Pos.X, .Pos.Y).Trigger = eTrigger.ZONASEGURA Then
            Call WriteConsoleMsg(UserIndex, "No puedes invocar criaturas en zona segura.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
    
        'No permitimos se invoquen criaturas en mapas donde esta prohibido hacerlo
        If MapInfo(Mapa).InvocarSinEfecto = 1 Then
            Call WriteConsoleMsg(UserIndex, "Invocar no esta permitido aqui! Retirate de la Zona si deseas utilizar el Hechizo.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
    
        Dim spellIndex As Integer, NroNpcs As Integer, NpcIndex As Integer, PetIndex As Integer

        Dim TargetPos  As WorldPos
    
        TargetPos.Map = .flags.TargetMap
        TargetPos.X = .flags.TargetX
        TargetPos.Y = .flags.TargetY
    
        spellIndex = .flags.Hechizo
    
        ' Warp de mascotas
        If Hechizos(spellIndex).Warp = 1 Then
            PetIndex = FarthestPet(UserIndex)
        
            ' La invoco cerca mio
            If PetIndex > 0 Then
                Call WarpMascota(UserIndex, PetIndex)

            End If
        
            ' Invocacion normal
        Else

            If .NroMascotas >= MAXMASCOTAS Then Exit Sub
        
            For NroNpcs = 1 To Hechizos(spellIndex).cant
            
                If .NroMascotas < MAXMASCOTAS Then
                    NpcIndex = SpawnNpc(Hechizos(spellIndex).NumNpc, TargetPos, True, False)

                    If NpcIndex > 0 Then
                        .NroMascotas = .NroMascotas + 1
                    
                        PetIndex = FreeMascotaIndex(UserIndex)
                    
                        .MascotasIndex(PetIndex) = NpcIndex
                        .MascotasType(PetIndex) = Npclist(NpcIndex).Numero
                    
                        With Npclist(NpcIndex)
                            .MaestroUser = UserIndex
                            .Contadores.TiempoExistencia = IntervaloInvocacion
                            .GiveGLD = 0

                        End With
                    
                        Call FollowAmo(NpcIndex)
                    Else
                        Exit Sub

                    End If

                Else
                    Exit For

                End If
        
            Next NroNpcs

        End If

    End With

    Call InfoHechizo(UserIndex)
    HechizoCasteado = True

    Exit Sub

Error:

    With UserList(UserIndex)
        LogError ("[" & Err.Number & "] " & Err.description & " por el usuario " & .Name & "(" & UserIndex & ") en (" & .Pos.Map & ", " & .Pos.X & ", " & .Pos.Y & "). Tratando de tirar el hechizo " & spellIndex & "(" & spellIndex & ") en la posicion ( " & .flags.TargetX & ", " & .flags.TargetY & ")")

    End With

End Sub

Sub HandleHechizoTerreno(ByVal UserIndex As Integer, ByVal spellIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: 18/11/2009
    '18/11/2009: ZaMa - Optimizacion de codigo.
    '***************************************************
    
    Dim HechizoCasteado As Boolean

    Dim ManaRequerida   As Integer
    
    Select Case Hechizos(spellIndex).Tipo

        Case TipoHechizo.uInvocacion
            Call HechizoInvocacion(UserIndex, HechizoCasteado)
            
        Case TipoHechizo.uEstado
            Call HechizoTerrenoEstado(UserIndex, HechizoCasteado)

    End Select

    If HechizoCasteado Then

        With UserList(UserIndex)
            Call SubirSkill(UserIndex, eSkill.Magia, True)
            
            ManaRequerida = Hechizos(spellIndex).ManaRequerido
            
            If Hechizos(spellIndex).Warp = 1 Then ' Invoco una mascota
                ' Consume toda la mana
                ManaRequerida = .Stats.MinMAN
            Else

                ' Bonificaciones en hechizos
                If .clase = eClass.Druid Then

                    ' Solo con flauta equipada
                    If .Invent.AnilloEqpObjIndex = FLAUTAELFICA Then
                        ' 30% menos de mana para invocaciones
                        ManaRequerida = ManaRequerida * 0.7

                    End If

                End If

            End If
            
            ' Quito la mana requerida
            .Stats.MinMAN = .Stats.MinMAN - ManaRequerida

            If .Stats.MinMAN < 0 Then .Stats.MinMAN = 0
            
            ' Quito la estamina requerida
            .Stats.MinSta = .Stats.MinSta - Hechizos(spellIndex).StaRequerido

            If .Stats.MinSta < 0 Then .Stats.MinSta = 0
            
            ' Update user stats
            Call WriteUpdateUserStats(UserIndex)

        End With

    End If
    
End Sub

Sub HandleHechizoUsuario(ByVal UserIndex As Integer, ByVal spellIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: 12/01/2010
    '18/11/2009: ZaMa - Optimizacion de codigo.
    '12/01/2010: ZaMa - Optimizacion y agrego bonificaciones al druida.
    '***************************************************
    
    Dim HechizoCasteado As Boolean

    Dim ManaRequerida   As Integer
    
    Select Case Hechizos(spellIndex).Tipo

        Case TipoHechizo.uEstado
            ' Afectan estados (por ejem : Envenenamiento)
            Call HechizoEstadoUsuario(UserIndex, HechizoCasteado)
        
        Case TipoHechizo.uPropiedades
            ' Afectan HP,MANA,STAMINA,ETC
            HechizoCasteado = HechizoPropUsuario(UserIndex)

    End Select

    If HechizoCasteado Then

        With UserList(UserIndex)
            Call SubirSkill(UserIndex, eSkill.Magia, True)
            
            ManaRequerida = Hechizos(spellIndex).ManaRequerido
            
            ' Bonificaciones para druida
            If .clase = eClass.Druid Then

                ' Solo con flauta magica
                If .Invent.AnilloEqpObjIndex = FLAUTAELFICA Then
                    If Hechizos(spellIndex).Mimetiza = 1 Then
                        ' 50% menos de mana para mimetismo
                        ManaRequerida = ManaRequerida * 0.5
                        
                    ElseIf spellIndex <> APOCALIPSIS_SPELL_INDEX Then
                        ' 10% menos de mana para todo menos apoca y descarga
                        ManaRequerida = ManaRequerida * 0.9

                    End If

                End If

            End If
            
            ' Quito la mana requerida
            .Stats.MinMAN = .Stats.MinMAN - ManaRequerida

            If .Stats.MinMAN < 0 Then .Stats.MinMAN = 0
            
            ' Quito la estamina requerida
            .Stats.MinSta = .Stats.MinSta - Hechizos(spellIndex).StaRequerido

            If .Stats.MinSta < 0 Then .Stats.MinSta = 0
            
            ' Update user stats
            Call WriteUpdateUserStats(UserIndex)
            Call WriteUpdateUserStats(.flags.TargetUser)
            .flags.TargetUser = 0

        End With

    End If

End Sub

Sub HandleHechizoNPC(ByVal UserIndex As Integer, ByVal HechizoIndex As Integer)

    '***************************************************
    'Author: Unknown
    'Last Modification: 12/01/2010
    '13/02/2009: ZaMa - Agregada 50% bonificacion en coste de mana a mimetismo para druidas
    '17/11/2009: ZaMa - Optimizacion de codigo.
    '12/01/2010: ZaMa - Bonificacion para druidas de 10% para todos hechizos excepto apoca y descarga.
    '12/01/2010: ZaMa - Los druidas mimetizados con npcs ahora son ignorados.
    '***************************************************
    Dim HechizoCasteado As Boolean

    Dim ManaRequerida   As Long
    
    With UserList(UserIndex)
        
        Select Case Hechizos(HechizoIndex).Tipo

            Case TipoHechizo.uEstado
                ' Afectan estados (por ejem : Envenenamiento)
                Call HechizoEstadoNPC(.flags.TargetNPC, HechizoIndex, HechizoCasteado, UserIndex)
                
            Case TipoHechizo.uPropiedades
                ' Afectan HP,MANA,STAMINA,ETC
                Call HechizoPropNPC(HechizoIndex, .flags.TargetNPC, UserIndex, HechizoCasteado)

        End Select
        
        If HechizoCasteado Then
            Call SubirSkill(UserIndex, eSkill.Magia, True)
            
            ManaRequerida = Hechizos(HechizoIndex).ManaRequerido
            
            ' Bonificacion para druidas.
            If .clase = eClass.Druid Then
                ' Se mostro como usuario, puede ser atacado por npcs
                .flags.Ignorado = False
                
                ' Solo con flauta equipada
                If .Invent.AnilloEqpObjIndex = FLAUTAELFICA Then
                    If Hechizos(HechizoIndex).Mimetiza = 1 Then
                        ' 50% menos de mana para mimetismo
                        ManaRequerida = ManaRequerida * 0.5
                        ' Sera ignorado hasta que pierda el efecto del mimetismo o ataque un npc
                        .flags.Ignorado = True
                    Else

                        ' 10% menos de mana para hechizos
                        If HechizoIndex <> APOCALIPSIS_SPELL_INDEX Then
                            ManaRequerida = ManaRequerida * 0.9

                        End If

                    End If

                End If

            End If
            
            ' Quito la mana requerida
            .Stats.MinMAN = .Stats.MinMAN - ManaRequerida

            If .Stats.MinMAN < 0 Then .Stats.MinMAN = 0
            
            ' Quito la estamina requerida
            .Stats.MinSta = .Stats.MinSta - Hechizos(HechizoIndex).StaRequerido

            If .Stats.MinSta < 0 Then .Stats.MinSta = 0
            
            ' Update user stats
            Call WriteUpdateUserStats(UserIndex)
            .flags.TargetNPC = 0

        End If

    End With

End Sub

Sub LanzarHechizo(ByVal spellIndex As Integer, ByVal UserIndex As Integer)

    '***************************************************
    'Autor: Unknown (orginal version)
    'Last Modification: 15/03/2020
    '24/01/2007 ZaMa - Optimizacion de codigo.
    '02/16/2010: Marco - Now .flags.hechizo makes reference to global spell index instead of user's spell index
    '15/03/2020: WyroX - Remuevo los chequeos de distancia, porque ya se comprueba si lanzo a un tile que ve
    '***************************************************
    On Error GoTo Errhandler

    With UserList(UserIndex)
    
        If .flags.EnConsulta Then
            Call WriteConsoleMsg(UserIndex, "No puedes lanzar hechizos si estas en consulta.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub

        End If
    
        If PuedeLanzar(UserIndex, spellIndex) Then

            Select Case Hechizos(spellIndex).Target

                Case TargetType.uUsuarios

                    If .flags.TargetUser > 0 Then
                        Call HandleHechizoUsuario(UserIndex, spellIndex)

                    Else
                        Call WriteConsoleMsg(UserIndex, "Este hechizo actua solo sobre usuarios.", FontTypeNames.FONTTYPE_INFO)

                    End If
            
                Case TargetType.uNPC

                    If .flags.TargetNPC > 0 Then
                        Call HandleHechizoNPC(UserIndex, spellIndex)

                    Else
                        Call WriteConsoleMsg(UserIndex, "Este hechizo solo afecta a los npcs.", FontTypeNames.FONTTYPE_INFO)

                    End If
            
                Case TargetType.uUsuariosYnpc

                    If .flags.TargetUser > 0 Then
                        Call HandleHechizoUsuario(UserIndex, spellIndex)

                    ElseIf .flags.TargetNPC > 0 Then
                        Call HandleHechizoNPC(UserIndex, spellIndex)

                    End If
            
                Case TargetType.uTerreno
                    Call HandleHechizoTerreno(UserIndex, spellIndex)

            End Select
        
        End If
    
        If .Counters.Trabajando Then .Counters.Trabajando = .Counters.Trabajando - 1
    
        If .Counters.Ocultando Then .Counters.Ocultando = .Counters.Ocultando - 1

    End With

    Exit Sub

Errhandler:
    Call LogError("Error en LanzarHechizo. Error " & Err.Number & " : " & Err.description & " Hechizo: " & spellIndex & "(" & spellIndex & "). Casteado por: " & UserList(UserIndex).Name & "(" & UserIndex & ").")
    
End Sub

Sub HechizoEstadoUsuario(ByVal UserIndex As Integer, ByRef HechizoCasteado As Boolean)
    '***************************************************
    'Autor: Unknown (orginal version)
    'Last Modification: 03/02/2020
    'Handles the Spells that afect the Stats of an User
    '24/01/2007 Pablo (ToxicWaste) - Invisibilidad no permitida en Mapas con InviSinEfecto
    '26/01/2007 Pablo (ToxicWaste) - Cambios que permiten mejor manejo de ataques en los rings.
    '26/01/2007 Pablo (ToxicWaste) - Revivir no permitido en Mapas con ResuSinEfecto
    '02/01/2008 Marcos (ByVal) - Curar Veneno no permitido en usuarios muertos.
    '06/28/2008 NicoNZ - Agregue que se le de valor al flag Inmovilizado.
    '17/11/2008: NicoNZ - Agregado para quitar la penalizacion de vida en el ring y cambio de ecuacion.
    '13/02/2009: ZaMa - Arreglada ecuacion para quitar vida tras resucitar en rings.
    '23/11/2009: ZaMa - Optimizacion de codigo.
    '28/04/2010: ZaMa - Agrego Restricciones para ciudas respecto al estado atacable.
    '16/09/2010: ZaMa - Solo se hace invi para los clientes si no esta navegando.
    '03/02/2020: WyroX - Anillos anti-efectos
    '***************************************************

    Dim HechizoIndex As Integer
    Dim targetIndex  As Integer

    With UserList(UserIndex)
        HechizoIndex = .flags.Hechizo
        targetIndex = .flags.TargetUser
    
        ' <-------- Agrega Invisibilidad ---------->
        If Hechizos(HechizoIndex).Invisibilidad = 1 Then
            If UserList(targetIndex).flags.Muerto = 1 Then
                Call WriteConsoleMsg(UserIndex, "El usuario esta muerto!", FontTypeNames.FONTTYPE_INFO)
                HechizoCasteado = False
                Exit Sub

            End If
        
            If UserList(targetIndex).Counters.Saliendo Then
                If UserIndex <> targetIndex Then
                    Call WriteConsoleMsg(UserIndex, "El hechizo no tiene efecto!", FontTypeNames.FONTTYPE_INFO)
                    HechizoCasteado = False
                    Exit Sub
                Else
                    Call WriteConsoleMsg(UserIndex, "No puedes hacerte invisible mientras te encuentras saliendo!", FontTypeNames.FONTTYPE_WARNING)
                    HechizoCasteado = False
                    Exit Sub

                End If

            End If
        
            'No usar invi mapas InviSinEfecto
            If MapInfo(UserList(targetIndex).Pos.Map).InviSinEfecto > 0 Then
                Call WriteConsoleMsg(UserIndex, "La invisibilidad no funciona aqui!", FontTypeNames.FONTTYPE_INFO)
                HechizoCasteado = False
                Exit Sub

            End If
            
            'Si sos user, no uses este hechizo con GMS.
            If Not EsGm(UserIndex) And EsGm(targetIndex) Then
                HechizoCasteado = False
                Exit Sub
            End If
            
            ' Chequea si el status permite ayudar al otro usuario
            HechizoCasteado = CanSupportUser(UserIndex, targetIndex, True)

            If Not HechizoCasteado Then Exit Sub

            UserList(targetIndex).flags.invisible = 1
        
            ' Solo se hace invi para los clientes si no esta navegando
            If UserList(targetIndex).flags.Navegando = 0 Then
                Call SetInvisible(targetIndex, UserList(targetIndex).Char.CharIndex, True)

            End If
        
            Call InfoHechizo(UserIndex)
            HechizoCasteado = True

        End If
    
        ' <-------- Agrega Mimetismo ---------->
        If Hechizos(HechizoIndex).Mimetiza = 1 Then
            If UserList(targetIndex).flags.Muerto = 1 Then
                Exit Sub

            End If
        
            If UserList(targetIndex).flags.Navegando = 1 Then
                Exit Sub

            End If

            If .flags.Navegando = 1 Then
                Exit Sub

            End If
        
            'Si sos user, no uses este hechizo con GMS.
            If EsGm(targetIndex) Then
                Call WriteConsoleMsg(UserIndex, "No puedes mimetizar a un Game Master.", FontTypeNames.FONTTYPE_FIGHT)
                HechizoCasteado = False
                Exit Sub
            End If
        
            If .flags.Mimetizado = 1 Then
                Call WriteConsoleMsg(UserIndex, "Ya te encuentras mimetizado. El hechizo no ha tenido efecto.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            If .flags.AdminInvisible = 1 Then Exit Sub
        
            'copio el char original al mimetizado
        
            .CharMimetizado.body = .Char.body
            .CharMimetizado.Head = .Char.Head
            .CharMimetizado.CascoAnim = .Char.CascoAnim
            .CharMimetizado.ShieldAnim = .Char.ShieldAnim
            .CharMimetizado.WeaponAnim = .Char.WeaponAnim
        
            .flags.Mimetizado = 1
        
            'ahora pongo local el del enemigo
            .Char.body = UserList(targetIndex).Char.body
            .Char.Head = UserList(targetIndex).Char.Head
            .Char.CascoAnim = UserList(targetIndex).Char.CascoAnim
            .Char.ShieldAnim = UserList(targetIndex).Char.ShieldAnim
            .Char.WeaponAnim = UserList(targetIndex).Char.WeaponAnim
        
            Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
       
            Call InfoHechizo(UserIndex)
            HechizoCasteado = True

        End If
    
        ' <-------- Agrega Envenenamiento ---------->
        If Hechizos(HechizoIndex).Envenena = 1 Then
            If UserIndex = targetIndex Then
                Call WriteConsoleMsg(UserIndex, "No puedes atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub

            End If
            
            'Si sos user, no uses este hechizo con GMS.
            If EsGm(targetIndex) Then
                Call WriteConsoleMsg(UserIndex, "Los Game Masters son inmunes a las alteraciones de estado.", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
            End If
            
            If Not PuedeAtacar(UserIndex, targetIndex) Then Exit Sub
            If UserIndex <> targetIndex Then
                Call UsuarioAtacadoPorUsuario(UserIndex, targetIndex)

            End If

            UserList(targetIndex).flags.Envenenado = 1
            Call InfoHechizo(UserIndex)
            HechizoCasteado = True

        End If
    
        ' <-------- Cura Envenenamiento ---------->
        If Hechizos(HechizoIndex).CuraVeneno = 1 Then
    
            'Verificamos que el usuario no este muerto
            If UserList(targetIndex).flags.Muerto = 1 Then
                Call WriteConsoleMsg(UserIndex, "El usuario esta muerto!", FontTypeNames.FONTTYPE_INFO)
                HechizoCasteado = False
                Exit Sub

            End If
            
            ' Chequea si el status permite ayudar al otro usuario
            HechizoCasteado = CanSupportUser(UserIndex, targetIndex)

            If Not HechizoCasteado Then Exit Sub
            
            UserList(targetIndex).flags.Envenenado = 0
            
            Call InfoHechizo(UserIndex)
            
            HechizoCasteado = True

        End If
    
        ' <-------- Agrega Maldicion ---------->
        If Hechizos(HechizoIndex).Maldicion = 1 Then
            If UserIndex = targetIndex Then
                Call WriteConsoleMsg(UserIndex, "No puedes atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub

            End If
            
            'Si sos user, no uses este hechizo con GMS.
            If EsGm(targetIndex) Then
                Call WriteConsoleMsg(UserIndex, "Los Game Masters son inmunes a las alteraciones de estado.", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
            End If
        
            If Not PuedeAtacar(UserIndex, targetIndex) Then Exit Sub
            
            If UserIndex <> targetIndex Then
                Call UsuarioAtacadoPorUsuario(UserIndex, targetIndex)

            End If

            UserList(targetIndex).flags.Maldicion = 1
            Call InfoHechizo(UserIndex)
            HechizoCasteado = True

        End If
    
        ' <-------- Remueve Maldicion ---------->
        If Hechizos(HechizoIndex).RemoverMaldicion = 1 Then
            UserList(targetIndex).flags.Maldicion = 0
            Call InfoHechizo(UserIndex)
            HechizoCasteado = True

        End If
    
        ' <-------- Agrega Bendicion ---------->
        If Hechizos(HechizoIndex).Bendicion = 1 Then
            UserList(targetIndex).flags.Bendicion = 1
            Call InfoHechizo(UserIndex)
            HechizoCasteado = True

        End If
    
        ' <-------- Agrega Paralisis/Inmobilidad ---------->
        If Hechizos(HechizoIndex).Paraliza = 1 Or Hechizos(HechizoIndex).Inmoviliza = 1 Then
            If UserIndex = targetIndex Then
                Call WriteConsoleMsg(UserIndex, "No puedes atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub

            End If
            
            'Si sos user, no uses este hechizo con GMS.
            If EsGm(targetIndex) Then
                Call WriteConsoleMsg(UserIndex, "Los Game Masters son inmunes a las alteraciones de estado.", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
            End If
            
            If UserList(targetIndex).flags.Paralizado = 0 Then
                If Not PuedeAtacar(UserIndex, targetIndex) Then Exit Sub
            
                If UserIndex <> targetIndex Then
                    Call UsuarioAtacadoPorUsuario(UserIndex, targetIndex)

                End If
            
                Call InfoHechizo(UserIndex)
                HechizoCasteado = True

                If UserList(targetIndex).Invent.AnilloEqpObjIndex > 0 Then
                    If ObjData(UserList(targetIndex).Invent.AnilloEqpObjIndex).ImpideParalizar Then
                        Call WriteConsoleMsg(targetIndex, "Tu anillo rechaza los efectos del hechizo.", FontTypeNames.FONTTYPE_FIGHT)
                        Call WriteConsoleMsg(UserIndex, "El hechizo no tiene efecto!", FontTypeNames.FONTTYPE_FIGHT)
                        Exit Sub
                    End If
                End If
            
                If Hechizos(HechizoIndex).Inmoviliza = 1 Then UserList(targetIndex).flags.Inmovilizado = 1
                UserList(targetIndex).flags.Paralizado = 1
                UserList(targetIndex).Counters.Paralisis = IntervaloParalizado
            
                UserList(targetIndex).flags.ParalizedByIndex = UserIndex
                UserList(targetIndex).flags.ParalizedBy = UserList(UserIndex).Name
            
                Call WriteParalizeOK(targetIndex)

            End If

        End If
    
        ' <-------- Remueve Paralisis/Inmobilidad ---------->
        If Hechizos(HechizoIndex).RemoverParalisis = 1 Then
        
            ' Remueve si esta en ese estado
            If UserList(targetIndex).flags.Paralizado = 1 Then
        
                ' Chequea si el status permite ayudar al otro usuario
                HechizoCasteado = CanSupportUser(UserIndex, targetIndex, True)

                If Not HechizoCasteado Then Exit Sub
            
                Call RemoveParalisis(targetIndex)
                Call InfoHechizo(UserIndex)
        
            End If

        End If
    
        ' <-------- Remueve Estupidez (Aturdimiento) ---------->
        If Hechizos(HechizoIndex).RemoverEstupidez = 1 Then
    
            ' Remueve si esta en ese estado
            If UserList(targetIndex).flags.Estupidez = 1 Then
        
                ' Chequea si el status permite ayudar al otro usuario
                HechizoCasteado = CanSupportUser(UserIndex, targetIndex)

                If Not HechizoCasteado Then Exit Sub
        
                UserList(targetIndex).flags.Estupidez = 0
            
                'no need to crypt this
                Call WriteDumbNoMore(targetIndex)
                Call InfoHechizo(UserIndex)
        
            End If

        End If
    
        ' <-------- Revive ---------->
        If Hechizos(HechizoIndex).Revivir = 1 Then
            If UserList(targetIndex).flags.Muerto = 1 Then
            
                'Seguro de resurreccion (solo afecta a los hechizos, no al sacerdote ni al comando de GM)
                If UserList(targetIndex).flags.SeguroResu Then
                    Call WriteConsoleMsg(UserIndex, "El espiritu no tiene intenciones de regresar al mundo de los vivos!", FontTypeNames.FONTTYPE_INFO)
                    HechizoCasteado = False
                    Exit Sub

                End If
        
                'No usar resu en mapas con ResuSinEfecto
                If MapInfo(UserList(targetIndex).Pos.Map).ResuSinEfecto > 0 Then
                    Call WriteConsoleMsg(UserIndex, "Revivir no esta permitido aqui! Retirate de la Zona si deseas utilizar el Hechizo.", FontTypeNames.FONTTYPE_INFO)
                    HechizoCasteado = False
                    Exit Sub

                End If
            
                'No podemos resucitar si nuestra barra de energia no esta llena. (GD: 29/04/07)
                If .Stats.MaxSta <> .Stats.MinSta Then
                    Call WriteConsoleMsg(UserIndex, "No puedes resucitar si no tienes tu barra de energia llena.", FontTypeNames.FONTTYPE_INFO)
                    HechizoCasteado = False
                    Exit Sub

                End If
            
                'revisamos si necesita vara
                If .clase = eClass.Mage Then
                    If .Invent.WeaponEqpObjIndex > 0 Then
                        If ObjData(.Invent.WeaponEqpObjIndex).StaffPower < Hechizos(HechizoIndex).NeedStaff Then
                            Call WriteConsoleMsg(UserIndex, "Necesitas un baculo mejor para lanzar este hechizo.", FontTypeNames.FONTTYPE_INFO)
                            HechizoCasteado = False
                            Exit Sub

                        End If

                    End If

                ElseIf .clase = eClass.Bard Then

                    If .Invent.AnilloEqpObjIndex <> LAUDELFICO And .Invent.AnilloEqpObjIndex <> LAUDMAGICO Then
                        Call WriteConsoleMsg(UserIndex, "Necesitas un instrumento magico para devolver la vida.", FontTypeNames.FONTTYPE_INFO)
                        HechizoCasteado = False
                        Exit Sub

                    End If

                ElseIf .clase = eClass.Druid Then

                    If .Invent.AnilloEqpObjIndex <> FLAUTAELFICA And .Invent.AnilloEqpObjIndex <> FLAUTAMAGICA Then
                        Call WriteConsoleMsg(UserIndex, "Necesitas un instrumento magico para devolver la vida.", FontTypeNames.FONTTYPE_INFO)
                        HechizoCasteado = False
                        Exit Sub

                    End If

                End If
            
                ' Chequea si el status permite ayudar al otro usuario
                HechizoCasteado = CanSupportUser(UserIndex, targetIndex, True)

                If Not HechizoCasteado Then Exit Sub
    
                Dim EraCriminal As Boolean

                EraCriminal = criminal(UserIndex)
            
                If Not criminal(targetIndex) Then
                    If targetIndex <> UserIndex Then
                        .Reputacion.NobleRep = .Reputacion.NobleRep + 500

                        If .Reputacion.NobleRep > MAXREP Then .Reputacion.NobleRep = MAXREP
                        Call WriteConsoleMsg(UserIndex, "Los Dioses te sonrien, has ganado 500 puntos de nobleza!", FontTypeNames.FONTTYPE_INFO)

                    End If

                End If
            
                If EraCriminal And Not criminal(UserIndex) Then
                    Call RefreshCharStatus(UserIndex)

                End If
            
                With UserList(targetIndex)
                    'Pablo Toxic Waste (GD: 29/04/07)
                    .Stats.MinAGU = 0
                    .flags.Sed = 1
                    .Stats.MinHam = 0
                    .flags.Hambre = 1
                    Call WriteUpdateHungerAndThirst(targetIndex)
                    Call InfoHechizo(UserIndex)
                    .Stats.MinMAN = 0
                    .Stats.MinSta = 0

                End With
            
                'Agregado para quitar la penalizacion de vida en el ring y cambio de ecuacion. (NicoNZ)
                If (TriggerZonaPelea(UserIndex, targetIndex) <> TRIGGER6_PERMITE) Then

                    'Solo saco vida si es User. no quiero que exploten GMs por ahi.
                    If .flags.Privilegios And PlayerType.User Then
                        .Stats.MinHp = .Stats.MinHp * (1 - UserList(targetIndex).Stats.ELV * 0.015)

                    End If

                End If
            
                If (.Stats.MinHp <= 0) Then
                    Call UserDie(UserIndex)
                    Call WriteConsoleMsg(UserIndex, "El esfuerzo de resucitar fue demasiado grande.", FontTypeNames.FONTTYPE_INFO)
                    HechizoCasteado = False
                Else
                    Call WriteConsoleMsg(UserIndex, "El esfuerzo de resucitar te ha debilitado.", FontTypeNames.FONTTYPE_INFO)
                    HechizoCasteado = True

                End If
            
                If UserList(targetIndex).flags.Traveling = 1 Then
                    UserList(targetIndex).Counters.goHome = 0
                    UserList(targetIndex).flags.Traveling = 0
                    'Call WriteConsoleMsg(TargetIndex, "Tu viaje ha sido cancelado.", FontTypeNames.FONTTYPE_FIGHT)
                    Call WriteMultiMessage(targetIndex, eMessages.CancelHome)

                End If
            
                Call RevivirUsuario(targetIndex)
            Else
                HechizoCasteado = False

            End If
    
        End If
    
        ' <-------- Agrega Ceguera ---------->
        If Hechizos(HechizoIndex).Ceguera = 1 Then
            If UserIndex = targetIndex Then
                Call WriteConsoleMsg(UserIndex, "No puedes atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub

            End If
            
            'Si sos user, no uses este hechizo con GMS.
            If EsGm(targetIndex) Then
                Call WriteConsoleMsg(UserIndex, "Los Game Masters son inmunes a las alteraciones de estado.", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
            End If
            
            If Not PuedeAtacar(UserIndex, targetIndex) Then Exit Sub

            If UserList(targetIndex).Invent.AnilloEqpObjIndex > 0 Then
                If ObjData(UserList(targetIndex).Invent.AnilloEqpObjIndex).ImpideCegar Then
                    Call WriteConsoleMsg(targetIndex, "Tu anillo rechaza los efectos del hechizo.", FontTypeNames.FONTTYPE_FIGHT)
                    Call WriteConsoleMsg(UserIndex, "El hechizo no tiene efecto!", FontTypeNames.FONTTYPE_FIGHT)
                    Exit Sub
                End If
            End If

            If UserIndex <> targetIndex Then
                Call UsuarioAtacadoPorUsuario(UserIndex, targetIndex)

            End If

            UserList(targetIndex).flags.Ceguera = 1
            UserList(targetIndex).Counters.Ceguera = IntervaloParalizado / 3
    
            Call WriteBlind(targetIndex)
            Call InfoHechizo(UserIndex)
            HechizoCasteado = True

        End If
    
        ' <-------- Agrega Estupidez (Aturdimiento) ---------->
        If Hechizos(HechizoIndex).Estupidez = 1 Then
            If UserIndex = targetIndex Then
                Call WriteConsoleMsg(UserIndex, "No puedes atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub

            End If
            
            'Si sos user, no uses este hechizo con GMS.
            If EsGm(targetIndex) Then
                Call WriteConsoleMsg(UserIndex, "Los Game Masters son inmunes a las alteraciones de estado.", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
            End If
            
            If Not PuedeAtacar(UserIndex, targetIndex) Then Exit Sub
                                                                                                                                
            If UserList(targetIndex).Invent.AnilloEqpObjIndex > 0 Then
                If ObjData(UserList(targetIndex).Invent.AnilloEqpObjIndex).ImpideAturdir Then
                    Call WriteConsoleMsg(targetIndex, "Tu anillo rechaza los efectos del hechizo.", FontTypeNames.FONTTYPE_FIGHT)
                    Call WriteConsoleMsg(UserIndex, "El hechizo no tiene efecto!", FontTypeNames.FONTTYPE_FIGHT)
                    Exit Sub
                End If
            End If

            If UserIndex <> targetIndex Then
                Call UsuarioAtacadoPorUsuario(UserIndex, targetIndex)

            End If

            If UserList(targetIndex).flags.Estupidez = 0 Then
                UserList(targetIndex).flags.Estupidez = 1
                UserList(targetIndex).Counters.Ceguera = IntervaloParalizado

            End If

            Call WriteDumb(targetIndex)
    
            Call InfoHechizo(UserIndex)
            HechizoCasteado = True

        End If

    End With

End Sub

Sub HechizoEstadoNPC(ByVal NpcIndex As Integer, _
                     ByVal spellIndex As Integer, _
                     ByRef HechizoCasteado As Boolean, _
                     ByVal UserIndex As Integer)
    '***************************************************
    'Autor: Unknown (orginal version)
    'Last Modification: 07/07/2008
    'Handles the Spells that afect the Stats of an NPC
    '04/13/2008 NicoNZ - Guardias Faccionarios pueden ser
    'removidos por users de su misma faccion.
    '07/07/2008: NicoNZ - Solo se puede mimetizar con npcs si es druida
    '***************************************************

    With Npclist(NpcIndex)

        If Hechizos(spellIndex).Invisibilidad = 1 Then
            Call InfoHechizo(UserIndex)
            .flags.invisible = 1
            HechizoCasteado = True

        End If
    
        If Hechizos(spellIndex).Envenena = 1 Then
            If Not PuedeAtacarNPC(UserIndex, NpcIndex) Then
                HechizoCasteado = False
                Exit Sub

            End If

            Call NPCAtacado(NpcIndex, UserIndex)
            Call InfoHechizo(UserIndex)
            .flags.Envenenado = 1
            HechizoCasteado = True

        End If
    
        If Hechizos(spellIndex).CuraVeneno = 1 Then
            Call InfoHechizo(UserIndex)
            .flags.Envenenado = 0
            HechizoCasteado = True

        End If
    
        If Hechizos(spellIndex).Maldicion = 1 Then
            If Not PuedeAtacarNPC(UserIndex, NpcIndex) Then
                HechizoCasteado = False
                Exit Sub

            End If

            Call NPCAtacado(NpcIndex, UserIndex)
            Call InfoHechizo(UserIndex)
            .flags.Maldicion = 1
            HechizoCasteado = True

        End If
    
        If Hechizos(spellIndex).RemoverMaldicion = 1 Then
            Call InfoHechizo(UserIndex)
            .flags.Maldicion = 0
            HechizoCasteado = True

        End If
    
        If Hechizos(spellIndex).Bendicion = 1 Then
            Call InfoHechizo(UserIndex)
            .flags.Bendicion = 1
            HechizoCasteado = True

        End If
    
        If Hechizos(spellIndex).Paraliza = 1 Then
            If .flags.AfectaParalisis = 0 Then
                If MapData(Npclist(NpcIndex).Pos.Map, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y).TileExit.Map > 0 Then
                    If Not EsGm(UserIndex) Then
                        Call WriteConsoleMsg(UserIndex, "No puedes paralizar criaturas en esa posicion.", FontTypeNames.FONTTYPE_INFOBOLD)   '"El NPC es inmune al hechizo."
                        HechizoCasteado = False
                        Exit Sub

                    End If

                End If
                                                                                                                      
                If Not PuedeAtacarNPC(UserIndex, NpcIndex, True) Then
                    HechizoCasteado = False
                    Exit Sub

                End If

                With UserList(UserIndex)
                    '<<<< Equitando >>>
                    If .flags.Equitando = 1 Then
                        Call UnmountMontura(UserIndex)
                        Call WriteEquitandoToggle(UserIndex)
                        
                    End If
                End With

                Call NPCAtacado(NpcIndex, UserIndex)
                Call InfoHechizo(UserIndex)
                .flags.Paralizado = 1
                .flags.Inmovilizado = 0
                .Contadores.Paralisis = IntervaloParalizado
                HechizoCasteado = True
            Else
                'Call WriteConsoleMsg(UserIndex, "El NPC es inmune a este hechizo.", FontTypeNames.FONTTYPE_INFO)
                Call WriteMultiMessage(UserIndex, eMessages.NpcInmune)
                HechizoCasteado = False
                Exit Sub

            End If

        End If
    
        If Hechizos(spellIndex).RemoverParalisis = 1 Then
            If .flags.Paralizado = 1 Or .flags.Inmovilizado = 1 Then
                If .MaestroUser = UserIndex Then
                    Call InfoHechizo(UserIndex)
                    .flags.Paralizado = 0
                    .Contadores.Paralisis = 0
                    HechizoCasteado = True
                Else

                    If .NPCtype = eNPCType.GuardiaReal Then
                        If esArmada(UserIndex) Then
                            Call InfoHechizo(UserIndex)
                            .flags.Paralizado = 0
                            .Contadores.Paralisis = 0
                            HechizoCasteado = True
                            Exit Sub
                        Else
                            Call WriteConsoleMsg(UserIndex, "Solo puedes remover la paralisis de los Guardias si perteneces a su faccion.", FontTypeNames.FONTTYPE_INFO)
                            HechizoCasteado = False
                            Exit Sub

                        End If
                    
                        Call WriteConsoleMsg(UserIndex, "Solo puedes remover la paralisis de los NPCs que te consideren su amo.", FontTypeNames.FONTTYPE_INFO)
                        HechizoCasteado = False
                        Exit Sub
                    Else

                        If .NPCtype = eNPCType.Guardiascaos Then
                            If esCaos(UserIndex) Then
                                Call InfoHechizo(UserIndex)
                                .flags.Paralizado = 0
                                .Contadores.Paralisis = 0
                                HechizoCasteado = True
                                Exit Sub
                            Else
                                Call WriteConsoleMsg(UserIndex, "Solo puedes remover la paralisis de los Guardias si perteneces a su faccion.", FontTypeNames.FONTTYPE_INFO)
                                HechizoCasteado = False
                                Exit Sub

                            End If

                        End If

                    End If

                End If

            Else
                Call WriteConsoleMsg(UserIndex, "Este NPC no esta paralizado", FontTypeNames.FONTTYPE_INFO)
                HechizoCasteado = False
                Exit Sub

            End If

        End If
     
        If Hechizos(spellIndex).Inmoviliza = 1 Then
            If .flags.AfectaParalisis = 0 Then
                If Not PuedeAtacarNPC(UserIndex, NpcIndex, True) Then
                    HechizoCasteado = False
                    Exit Sub

                End If

                With UserList(UserIndex)
                '<<<< Equitando >>>
                    If .flags.Equitando = 1 Then
                        Call UnmountMontura(UserIndex)
                        Call WriteEquitandoToggle(UserIndex)
                        
                    End If
                End With

                If MapData(Npclist(NpcIndex).Pos.Map, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y).TileExit.Map > 0 Then
                    If Not EsGm(UserIndex) Then
                        Call WriteConsoleMsg(UserIndex, "No puedes paralizar criaturas en esa posicion.", FontTypeNames.FONTTYPE_INFOBOLD)   '"El NPC es inmune al hechizo."
                        HechizoCasteado = False
                        Exit Sub

                    End If

                End If
                                                                                                                                            
                Call NPCAtacado(NpcIndex, UserIndex)
                .flags.Inmovilizado = 1
                .flags.Paralizado = 0
                .Contadores.Paralisis = IntervaloParalizado
                Call InfoHechizo(UserIndex)
                HechizoCasteado = True
            Else
                'Call WriteConsoleMsg(UserIndex, "El NPC es inmune al hechizo.", FontTypeNames.FONTTYPE_INFO)
                Call WriteMultiMessage(UserIndex, eMessages.NpcInmune)

            End If

        End If

    End With

    If Hechizos(spellIndex).Mimetiza = 1 Then

        With UserList(UserIndex)

            If .flags.Mimetizado = 1 Then
                Call WriteConsoleMsg(UserIndex, "Ya te encuentras mimetizado. El hechizo no ha tenido efecto.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
        
            If .flags.AdminInvisible = 1 Then Exit Sub
            
            If .clase = eClass.Druid Then
                'copio el char original al mimetizado
            
                .CharMimetizado.body = .Char.body
                .CharMimetizado.Head = .Char.Head
                .CharMimetizado.CascoAnim = .Char.CascoAnim
                .CharMimetizado.ShieldAnim = .Char.ShieldAnim
                .CharMimetizado.WeaponAnim = .Char.WeaponAnim
            
                .flags.Mimetizado = 1
            
                'ahora pongo lo del NPC.
                .Char.body = Npclist(NpcIndex).Char.body
                .Char.Head = Npclist(NpcIndex).Char.Head
                .Char.CascoAnim = NingunCasco
                .Char.ShieldAnim = NingunEscudo
                .Char.WeaponAnim = NingunArma
        
                Call ChangeUserChar(UserIndex, .Char.body, .Char.Head, .Char.heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
            
            Else
                Call WriteConsoleMsg(UserIndex, "Solo los druidas pueden mimetizarse con criaturas.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub

            End If
    
            Call InfoHechizo(UserIndex)
            HechizoCasteado = True

        End With

    End If

End Sub

Sub HechizoPropNPC(ByVal spellIndex As Integer, _
                   ByVal NpcIndex As Integer, _
                   ByVal UserIndex As Integer, _
                   ByRef HechizoCasteado As Boolean)
    '***************************************************
    'Autor: Unknown (orginal version)
    'Last Modification: 06/04/2020
    'Handles the Spells that afect the Life NPC
    '14/08/2007 Pablo (ToxicWaste) - Orden general.
    '18/09/2010: ZaMa - Ahora valida si podes ayudar a un npc.
    '06/04/2020: FrankoH298 - Si le lanza un hechizo al npc lo desmonta.
    '***************************************************

    Dim dano As Long

    With Npclist(NpcIndex)
    
        Dim tempX, tempY As Integer

        tempX = .Pos.X
        tempY = .Pos.Y
        'Salud
        If Hechizos(spellIndex).SubeHP = 1 Then
        
            HechizoCasteado = CanSupportNpc(UserIndex, NpcIndex)
        
            If HechizoCasteado Then
                dano = RandomNumber(Hechizos(spellIndex).MinHp, Hechizos(spellIndex).MaxHp)
                dano = dano + Porcentaje(dano, 3 * UserList(UserIndex).Stats.ELV)
            
                Call InfoHechizo(UserIndex)
                .Stats.MinHp = .Stats.MinHp + dano

                If .Stats.MinHp > .Stats.MaxHp Then .Stats.MinHp = .Stats.MaxHp
                Call WriteConsoleMsg(UserIndex, "Has curado " & dano & " puntos de vida a la criatura.", FontTypeNames.FONTTYPE_FIGHT)
                
                'Renderizo el dano en render
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateDamage(.Pos.X, .Pos.Y, dano, DAMAGE_CURAR))
                
            End If
        
        ElseIf Hechizos(spellIndex).SubeHP = 2 Then

            If Not PuedeAtacarNPC(UserIndex, NpcIndex) Then
                HechizoCasteado = False
                Exit Sub

            End If

            With UserList(UserIndex)
                '<<<< Equitando >>>
                If .flags.Equitando = 1 Then
                    Call UnmountMontura(UserIndex)
                    Call WriteEquitandoToggle(UserIndex)
                    
                End If
            End With

            Call NPCAtacado(NpcIndex, UserIndex)
            dano = RandomNumber(Hechizos(spellIndex).MinHp, Hechizos(spellIndex).MaxHp)
            dano = dano + Porcentaje(dano, 3 * UserList(UserIndex).Stats.ELV)
    
            If Hechizos(spellIndex).StaffAffected Then
                If UserList(UserIndex).clase = eClass.Mage Then
                    If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                        dano = (dano * (ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).StaffDamageBonus + 70)) / 100
                        'Aumenta dano segun el staff-
                        'Dano = (Dano* (70 + BonifBaculo)) / 100
                    Else
                        dano = dano * 0.7 'Baja dano a 70% del original

                    End If

                End If

            End If

            If UserList(UserIndex).Invent.AnilloEqpObjIndex = LAUDELFICO Then
                dano = dano * 1.02  'Laud magico de los bardos
            ElseIf UserList(UserIndex).Invent.AnilloEqpObjIndex = FLAUTAELFICA Then
                dano = dano * 1.04 'Flauta
            End If
    
            Call InfoHechizo(UserIndex)
            HechizoCasteado = True
        
            If .flags.Snd2 > 0 Then
                Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessagePlayWave(.flags.Snd2, .Pos.X, .Pos.Y))

            End If
        
            'Quizas tenga defenza magica el NPC. Pablo (ToxicWaste)
            dano = dano - .Stats.defM

            If dano < 0 Then dano = 0
        
            Call EventosDano(UserIndex, NpcIndex, dano)
        
            .Stats.MinHp = .Stats.MinHp - dano
            Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCreateDamage(.Pos.X, .Pos.Y, dano, DAMAGE_NORMAL))
            'Call WriteConsoleMsg(UserIndex, "Le has quitado " & dano & " puntos de vida a la criatura!", FontTypeNames.FONTTYPE_FIGHT)
            Call WriteMultiMessage(UserIndex, eMessages.UserHitNPC, dano)
            Call CalcularDarExp(UserIndex, NpcIndex, dano)
    
            If .Stats.MinHp < 1 Then
                .Stats.MinHp = 0
                Call MuereNpc(NpcIndex, UserIndex)
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageFXtoMap(Hechizos(spellIndex).FXgrh, Hechizos(spellIndex).loops, tempX, tempY))

            End If

        End If

    End With

End Sub

Sub InfoHechizo(ByVal UserIndex As Integer)

    '***************************************************
    'Autor: Unknown (orginal version)
    'Last Modification: 25/07/2009
    '25/07/2009: ZaMa - Code improvements.
    '25/07/2009: ZaMa - Now invisible admins magic sounds are not sent to anyone but themselves
    '***************************************************
    Dim spellIndex As Integer
    Dim tUser      As Integer
    Dim tNPC       As Integer
    Dim tempData   As String
    
    With UserList(UserIndex)
        spellIndex = .flags.Hechizo
        tUser = .flags.TargetUser
        tNPC = .flags.TargetNPC
     
        Call DecirPalabrasMagicas(spellIndex, UserIndex)
     
        If tUser > 0 Then

            ' Los admins invisibles no producen sonidos ni fx's
            If .flags.AdminInvisible = 1 And UserIndex = tUser Then
                
                tempData = PrepareMessageCreateFX(UserList(tUser).Char.CharIndex, Hechizos(spellIndex).FXgrh, Hechizos(spellIndex).loops)
                Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(tempData)
                
                tempData = PrepareMessagePlayWave(Hechizos(spellIndex).WAV, UserList(tUser).Pos.X, UserList(tUser).Pos.Y)
                Call UserList(UserIndex).outgoingData.WriteASCIIStringFixed(tempData)

            Else
                Call SendData(SendTarget.ToPCArea, tUser, PrepareMessageCreateFX(UserList(tUser).Char.CharIndex, Hechizos(spellIndex).FXgrh, Hechizos(spellIndex).loops))
                Call SendData(SendTarget.ToPCArea, tUser, PrepareMessagePlayWave(Hechizos(spellIndex).WAV, UserList(tUser).Pos.X, UserList(tUser).Pos.Y)) 'Esta linea faltaba. Pablo (ToxicWaste)

            End If

        ElseIf tNPC > 0 Then
            Call SendData(SendTarget.ToNPCArea, tNPC, PrepareMessageCreateFX(Npclist(tNPC).Char.CharIndex, Hechizos(spellIndex).FXgrh, Hechizos(spellIndex).loops))
            Call SendData(SendTarget.ToNPCArea, tNPC, PrepareMessagePlayWave(Hechizos(spellIndex).WAV, Npclist(tNPC).Pos.X, Npclist(tNPC).Pos.Y))

        End If
     
        If tUser > 0 Then
            If UserIndex <> tUser Then
                If .showName Then
                    Call WriteMultiMessage(UserIndex, eMessages.Hechizo_HechiceroMSG_NOMBRE, spellIndex, , , UserList(tUser).Name)
                Else
                    Call WriteMultiMessage(UserIndex, eMessages.Hechizo_HechiceroMSG_ALGUIEN, spellIndex)

                End If

                Call WriteMultiMessage(tUser, eMessages.Hechizo_TargetMSG, spellIndex, , , .Name)
            Else
                Call WriteMultiMessage(UserIndex, eMessages.Hechizo_PropioMSG, spellIndex)

            End If

        ElseIf tNPC > 0 Then
            Call WriteMultiMessage(UserIndex, eMessages.Hechizo_HechiceroMSG_CRIATURA, spellIndex)

        End If

    End With
 
End Sub

Public Function HechizoPropUsuario(ByVal UserIndex As Integer) As Boolean
    '***************************************************
    'Autor: Unknown (orginal version)
    'Last Modification: 06/04/2020
    '02/01/2008 Marcos (ByVal) - No permite tirar curar heridas a usuarios muertos.
    '28/04/2010: ZaMa - Agrego Restricciones para ciudas respecto al estado atacable.
    '06/04/2020: FrankoH298 - Si le lanza un hechizo a un usuario lo desmonta.
    '***************************************************

    Dim spellIndex  As Integer

    Dim dano        As Long

    Dim targetIndex As Integer

    spellIndex = UserList(UserIndex).flags.Hechizo
    targetIndex = UserList(UserIndex).flags.TargetUser
      
    With UserList(targetIndex)

        If .flags.Muerto Then
            'Call WriteConsoleMsg(UserIndex, "No puedes lanzar este hechizo a un muerto.", FontTypeNames.FONTTYPE_INFO)
            Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
            Exit Function

        End If
        
        '<<<< Equitando >>>
        If .flags.Equitando = 1 Then
            Call UnmountMontura(targetIndex)
            Call WriteEquitandoToggle(targetIndex)
            
        End If

        ' <-------- Aumenta Hambre ---------->
        If Hechizos(spellIndex).SubeHam = 1 Then
        
            Call InfoHechizo(UserIndex)
        
            dano = RandomNumber(Hechizos(spellIndex).MinHam, Hechizos(spellIndex).MaxHam)
        
            .Stats.MinHam = .Stats.MinHam + dano

            If .Stats.MinHam > .Stats.MaxHam Then .Stats.MinHam = .Stats.MaxHam
        
            If UserIndex <> targetIndex Then
                Call WriteConsoleMsg(UserIndex, "Le has restaurado " & dano & " puntos de hambre a " & .Name & ".", FontTypeNames.FONTTYPE_FIGHT)
                Call WriteConsoleMsg(targetIndex, UserList(UserIndex).Name & " te ha restaurado " & dano & " puntos de hambre.", FontTypeNames.FONTTYPE_FIGHT)
            Else
                Call WriteConsoleMsg(UserIndex, "Te has restaurado " & dano & " puntos de hambre.", FontTypeNames.FONTTYPE_FIGHT)

            End If
        
            Call WriteUpdateHungerAndThirst(targetIndex)
    
            ' <-------- Quita Hambre ---------->
        ElseIf Hechizos(spellIndex).SubeHam = 2 Then

            If Not PuedeAtacar(UserIndex, targetIndex) Then Exit Function
        
            If UserIndex <> targetIndex Then
                Call UsuarioAtacadoPorUsuario(UserIndex, targetIndex)
            Else
                Exit Function

            End If
        
            Call InfoHechizo(UserIndex)
        
            dano = RandomNumber(Hechizos(spellIndex).MinHam, Hechizos(spellIndex).MaxHam)
        
            .Stats.MinHam = .Stats.MinHam - dano
        
            If UserIndex <> targetIndex Then
                Call WriteConsoleMsg(UserIndex, "Le has quitado " & dano & " puntos de hambre a " & .Name & ".", FontTypeNames.FONTTYPE_FIGHT)
                Call WriteConsoleMsg(targetIndex, UserList(UserIndex).Name & " te ha quitado " & dano & " puntos de hambre.", FontTypeNames.FONTTYPE_FIGHT)
            Else
                Call WriteConsoleMsg(UserIndex, "Te has quitado " & dano & " puntos de hambre.", FontTypeNames.FONTTYPE_FIGHT)

            End If
        
            If .Stats.MinHam < 1 Then
                .Stats.MinHam = 0
                .flags.Hambre = 1

            End If
        
            Call WriteUpdateHungerAndThirst(targetIndex)

        End If
    
        ' <-------- Aumenta Sed ---------->
        If Hechizos(spellIndex).SubeSed = 1 Then
        
            Call InfoHechizo(UserIndex)
        
            dano = RandomNumber(Hechizos(spellIndex).MinSed, Hechizos(spellIndex).MaxSed)
        
            .Stats.MinAGU = .Stats.MinAGU + dano

            If .Stats.MinAGU > .Stats.MaxAGU Then .Stats.MinAGU = .Stats.MaxAGU
        
            Call WriteUpdateHungerAndThirst(targetIndex)
             
            If UserIndex <> targetIndex Then
                Call WriteConsoleMsg(UserIndex, "Le has restaurado " & dano & " puntos de sed a " & .Name & ".", FontTypeNames.FONTTYPE_FIGHT)
                Call WriteConsoleMsg(targetIndex, UserList(UserIndex).Name & " te ha restaurado " & dano & " puntos de sed.", FontTypeNames.FONTTYPE_FIGHT)
            Else
                Call WriteConsoleMsg(UserIndex, "Te has restaurado " & dano & " puntos de sed.", FontTypeNames.FONTTYPE_FIGHT)

            End If
    
            ' <-------- Quita Sed ---------->
        ElseIf Hechizos(spellIndex).SubeSed = 2 Then
        
            If Not PuedeAtacar(UserIndex, targetIndex) Then Exit Function
        
            If UserIndex <> targetIndex Then
                Call UsuarioAtacadoPorUsuario(UserIndex, targetIndex)

            End If
        
            Call InfoHechizo(UserIndex)
        
            dano = RandomNumber(Hechizos(spellIndex).MinSed, Hechizos(spellIndex).MaxSed)
        
            .Stats.MinAGU = .Stats.MinAGU - dano
        
            If UserIndex <> targetIndex Then
                Call WriteConsoleMsg(UserIndex, "Le has quitado " & dano & " puntos de sed a " & .Name & ".", FontTypeNames.FONTTYPE_FIGHT)
                Call WriteConsoleMsg(targetIndex, UserList(UserIndex).Name & " te ha quitado " & dano & " puntos de sed.", FontTypeNames.FONTTYPE_FIGHT)
            Else
                Call WriteConsoleMsg(UserIndex, "Te has quitado " & dano & " puntos de sed.", FontTypeNames.FONTTYPE_FIGHT)

            End If
        
            If .Stats.MinAGU < 1 Then
                .Stats.MinAGU = 0
                .flags.Sed = 1

            End If
        
            Call WriteUpdateHungerAndThirst(targetIndex)
        
        End If
    
        ' <-------- Aumenta Agilidad ---------->
        If Hechizos(spellIndex).SubeAgilidad = 1 Then
        
            ' Chequea si el status permite ayudar al otro usuario
            If Not CanSupportUser(UserIndex, targetIndex) Then Exit Function
        
            Call InfoHechizo(UserIndex)
            dano = RandomNumber(Hechizos(spellIndex).MinAgilidad, Hechizos(spellIndex).MaxAgilidad)
        
            .flags.DuracionEfecto = 1200
            .Stats.UserAtributos(eAtributos.Agilidad) = .Stats.UserAtributos(eAtributos.Agilidad) + dano

            If .Stats.UserAtributos(eAtributos.Agilidad) > MinimoInt(MAXATRIBUTOS, .Stats.UserAtributosBackUP(Agilidad) * 2) Then .Stats.UserAtributos(eAtributos.Agilidad) = MinimoInt(MAXATRIBUTOS, .Stats.UserAtributosBackUP(Agilidad) * 2)
        
            .flags.TomoPocion = True
            Call WriteUpdateDexterity(targetIndex)
    
            ' <-------- Quita Agilidad ---------->
        ElseIf Hechizos(spellIndex).SubeAgilidad = 2 Then
        
            If Not PuedeAtacar(UserIndex, targetIndex) Then Exit Function
        
            If UserIndex <> targetIndex Then
                Call UsuarioAtacadoPorUsuario(UserIndex, targetIndex)

            End If
        
            Call InfoHechizo(UserIndex)
        
            .flags.TomoPocion = True
            dano = RandomNumber(Hechizos(spellIndex).MinAgilidad, Hechizos(spellIndex).MaxAgilidad)
            .flags.DuracionEfecto = 700
            .Stats.UserAtributos(eAtributos.Agilidad) = .Stats.UserAtributos(eAtributos.Agilidad) - dano

            If .Stats.UserAtributos(eAtributos.Agilidad) < MINATRIBUTOS Then .Stats.UserAtributos(eAtributos.Agilidad) = MINATRIBUTOS
        
            Call WriteUpdateDexterity(targetIndex)

        End If
    
        ' <-------- Aumenta Fuerza ---------->
        If Hechizos(spellIndex).SubeFuerza = 1 Then
    
            ' Chequea si el status permite ayudar al otro usuario
            If Not CanSupportUser(UserIndex, targetIndex) Then Exit Function
        
            Call InfoHechizo(UserIndex)
            dano = RandomNumber(Hechizos(spellIndex).MinFuerza, Hechizos(spellIndex).MaxFuerza)
        
            .flags.DuracionEfecto = 1200
    
            .Stats.UserAtributos(eAtributos.Fuerza) = .Stats.UserAtributos(eAtributos.Fuerza) + dano

            If .Stats.UserAtributos(eAtributos.Fuerza) > MinimoInt(MAXATRIBUTOS, .Stats.UserAtributosBackUP(Fuerza) * 2) Then .Stats.UserAtributos(eAtributos.Fuerza) = MinimoInt(MAXATRIBUTOS, .Stats.UserAtributosBackUP(Fuerza) * 2)
        
            .flags.TomoPocion = True
            Call WriteUpdateStrenght(targetIndex)
    
            ' <-------- Quita Fuerza ---------->
        ElseIf Hechizos(spellIndex).SubeFuerza = 2 Then
    
            If Not PuedeAtacar(UserIndex, targetIndex) Then Exit Function
        
            If UserIndex <> targetIndex Then
                Call UsuarioAtacadoPorUsuario(UserIndex, targetIndex)

            End If
        
            Call InfoHechizo(UserIndex)
        
            .flags.TomoPocion = True
        
            dano = RandomNumber(Hechizos(spellIndex).MinFuerza, Hechizos(spellIndex).MaxFuerza)
            .flags.DuracionEfecto = 700
            .Stats.UserAtributos(eAtributos.Fuerza) = .Stats.UserAtributos(eAtributos.Fuerza) - dano

            If .Stats.UserAtributos(eAtributos.Fuerza) < MINATRIBUTOS Then .Stats.UserAtributos(eAtributos.Fuerza) = MINATRIBUTOS
        
            Call WriteUpdateStrenght(targetIndex)

        End If
    
        ' <-------- Cura salud ---------->
        If Hechizos(spellIndex).SubeHP = 1 Then
        
            'Verifica que el usuario no este muerto
            If .flags.Muerto = 1 Then
                Call WriteConsoleMsg(UserIndex, "El usuario esta muerto!", FontTypeNames.FONTTYPE_INFO)
                Exit Function

            End If
        
            ' Chequea si el status permite ayudar al otro usuario
            If Not CanSupportUser(UserIndex, targetIndex) Then Exit Function
           
            dano = RandomNumber(Hechizos(spellIndex).MinHp, Hechizos(spellIndex).MaxHp)
            dano = dano + Porcentaje(dano, 3 * UserList(UserIndex).Stats.ELV)
        
            Call InfoHechizo(UserIndex)
    
            .Stats.MinHp = .Stats.MinHp + dano

            If .Stats.MinHp > .Stats.MaxHp Then .Stats.MinHp = .Stats.MaxHp
        
            Call WriteUpdateHP(targetIndex)
        
            If UserIndex <> targetIndex Then
                Call WriteConsoleMsg(UserIndex, "Le has restaurado " & dano & " puntos de vida a " & .Name & ".", FontTypeNames.FONTTYPE_FIGHT)
                Call WriteConsoleMsg(targetIndex, UserList(UserIndex).Name & " te ha restaurado " & dano & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
                
                'Renderizo el dano en render
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateDamage(.Pos.X, .Pos.Y, dano, DAMAGE_CURAR))
                Call SendData(SendTarget.ToPCArea, targetIndex, PrepareMessageCreateDamage(.Pos.X, .Pos.Y, dano, DAMAGE_CURAR))
                
            Else
                Call WriteConsoleMsg(UserIndex, "Te has restaurado " & dano & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
                
                'Renderizo el dano en render
                Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateDamage(.Pos.X, .Pos.Y, dano, DAMAGE_CURAR))

            End If
        
            ' <-------- Quita salud (Dana) ---------->
        ElseIf Hechizos(spellIndex).SubeHP = 2 Then
        
            If UserIndex = targetIndex Then
                Call WriteConsoleMsg(UserIndex, "No puedes atacarte a vos mismo.", FontTypeNames.FONTTYPE_FIGHT)
                Exit Function

            End If
        
            dano = RandomNumber(Hechizos(spellIndex).MinHp, Hechizos(spellIndex).MaxHp)
        
            dano = dano + Porcentaje(dano, 3 * UserList(UserIndex).Stats.ELV)
        
            If Hechizos(spellIndex).StaffAffected Then
                If UserList(UserIndex).clase = eClass.Mage Then
                    If UserList(UserIndex).Invent.WeaponEqpObjIndex > 0 Then
                        dano = (dano * (ObjData(UserList(UserIndex).Invent.WeaponEqpObjIndex).StaffDamageBonus + 70)) / 100
                    Else
                        dano = dano * 0.7 'Baja dano a 70% del original

                    End If

                End If

            End If
        
            If UserList(UserIndex).Invent.AnilloEqpObjIndex = LAUDELFICO Or UserList(UserIndex).Invent.AnilloEqpObjIndex = FLAUTAELFICA Then
                dano = dano * 1.04  'laud magico de los bardos

            End If
        
            'cascos antimagia
            If (.Invent.CascoEqpObjIndex > 0) Then
                dano = dano - RandomNumber(ObjData(.Invent.CascoEqpObjIndex).DefensaMagicaMin, ObjData(.Invent.CascoEqpObjIndex).DefensaMagicaMax)

            End If
        
            'anillos
            If (.Invent.AnilloEqpObjIndex > 0) Then
                dano = dano - RandomNumber(ObjData(.Invent.AnilloEqpObjIndex).DefensaMagicaMin, ObjData(.Invent.AnilloEqpObjIndex).DefensaMagicaMax)

            End If
        
            If dano < 0 Then dano = 0
        
            If Not PuedeAtacar(UserIndex, targetIndex) Then Exit Function
        
            If UserIndex <> targetIndex Then
                Call UsuarioAtacadoPorUsuario(UserIndex, targetIndex)

            End If
        
            Call InfoHechizo(UserIndex)
        
            .Stats.MinHp = .Stats.MinHp - dano
            
            'Renderizo el dano en render
            Call SendData(SendTarget.ToPCArea, targetIndex, PrepareMessageCreateDamage(.Pos.X, .Pos.Y, dano, DAMAGE_NORMAL))
            
            Call WriteUpdateHP(targetIndex)
        
            Call WriteConsoleMsg(UserIndex, "Le has quitado " & dano & " puntos de vida a " & .Name & ".", FontTypeNames.FONTTYPE_FIGHT)
            Call WriteConsoleMsg(targetIndex, UserList(UserIndex).Name & " te ha quitado " & dano & " puntos de vida.", FontTypeNames.FONTTYPE_FIGHT)
        
            'Muere
            If .Stats.MinHp < 1 Then
        
                If .flags.AtacablePor <> UserIndex Then
                    'Store it!
                    Call Statistics.StoreFrag(UserIndex, targetIndex)
                    Call ContarMuerte(targetIndex, UserIndex)

                End If
            
                .Stats.MinHp = 0
                Call ActStats(targetIndex, UserIndex)
                Call UserDie(targetIndex, UserIndex)

            End If
        
        End If
    
        ' <-------- Aumenta Mana ---------->
        If Hechizos(spellIndex).SubeMana = 1 Then
        
            Call InfoHechizo(UserIndex)
            .Stats.MinMAN = .Stats.MinMAN + dano

            If .Stats.MinMAN > .Stats.MaxMAN Then .Stats.MinMAN = .Stats.MaxMAN
        
            Call WriteUpdateMana(targetIndex)
        
            If UserIndex <> targetIndex Then
                Call WriteConsoleMsg(UserIndex, "Le has restaurado " & dano & " puntos de mana a " & .Name & ".", FontTypeNames.FONTTYPE_FIGHT)
                Call WriteConsoleMsg(targetIndex, UserList(UserIndex).Name & " te ha restaurado " & dano & " puntos de mana.", FontTypeNames.FONTTYPE_FIGHT)
            Else
                Call WriteConsoleMsg(UserIndex, "Te has restaurado " & dano & " puntos de mana.", FontTypeNames.FONTTYPE_FIGHT)

            End If
    
            ' <-------- Quita Mana ---------->
        ElseIf Hechizos(spellIndex).SubeMana = 2 Then

            If Not PuedeAtacar(UserIndex, targetIndex) Then Exit Function
        
            If UserIndex <> targetIndex Then
                Call UsuarioAtacadoPorUsuario(UserIndex, targetIndex)

            End If
        
            Call InfoHechizo(UserIndex)
        
            If UserIndex <> targetIndex Then
                Call WriteConsoleMsg(UserIndex, "Le has quitado " & dano & " puntos de mana a " & .Name & ".", FontTypeNames.FONTTYPE_FIGHT)
                Call WriteConsoleMsg(targetIndex, UserList(UserIndex).Name & " te ha quitado " & dano & " puntos de mana.", FontTypeNames.FONTTYPE_FIGHT)
            Else
                Call WriteConsoleMsg(UserIndex, "Te has quitado " & dano & " puntos de mana.", FontTypeNames.FONTTYPE_FIGHT)

            End If
        
            .Stats.MinMAN = .Stats.MinMAN - dano

            If .Stats.MinMAN < 1 Then .Stats.MinMAN = 0
        
            Call WriteUpdateMana(targetIndex)
        
        End If
    
        ' <-------- Aumenta Stamina ---------->
        If Hechizos(spellIndex).SubeSta = 1 Then
            Call InfoHechizo(UserIndex)
            .Stats.MinSta = .Stats.MinSta + dano

            If .Stats.MinSta > .Stats.MaxSta Then .Stats.MinSta = .Stats.MaxSta
        
            Call WriteUpdateSta(targetIndex)
        
            If UserIndex <> targetIndex Then
                Call WriteConsoleMsg(UserIndex, "Le has restaurado " & dano & " puntos de energia a " & .Name & ".", FontTypeNames.FONTTYPE_FIGHT)
                Call WriteConsoleMsg(targetIndex, UserList(UserIndex).Name & " te ha restaurado " & dano & " puntos de energia.", FontTypeNames.FONTTYPE_FIGHT)
            Else
                Call WriteConsoleMsg(UserIndex, "Te has restaurado " & dano & " puntos de energia.", FontTypeNames.FONTTYPE_FIGHT)

            End If
        
            ' <-------- Quita Stamina ---------->
        ElseIf Hechizos(spellIndex).SubeSta = 2 Then

            If Not PuedeAtacar(UserIndex, targetIndex) Then Exit Function
        
            If UserIndex <> targetIndex Then
                Call UsuarioAtacadoPorUsuario(UserIndex, targetIndex)

            End If
        
            Call InfoHechizo(UserIndex)
        
            If UserIndex <> targetIndex Then
                Call WriteConsoleMsg(UserIndex, "Le has quitado " & dano & " puntos de energia a " & .Name & ".", FontTypeNames.FONTTYPE_FIGHT)
                Call WriteConsoleMsg(targetIndex, UserList(UserIndex).Name & " te ha quitado " & dano & " puntos de energia.", FontTypeNames.FONTTYPE_FIGHT)
            Else
                Call WriteConsoleMsg(UserIndex, "Te has quitado " & dano & " puntos de energia.", FontTypeNames.FONTTYPE_FIGHT)

            End If
        
            .Stats.MinSta = .Stats.MinSta - dano
        
            If .Stats.MinSta < 1 Then .Stats.MinSta = 0
        
            Call WriteUpdateSta(targetIndex)
        
        End If

    End With

    HechizoPropUsuario = True

End Function

Public Function CanSupportUser(ByVal CasterIndex As Integer, _
                               ByVal targetIndex As Integer, _
                               Optional ByVal DoCriminal As Boolean = False) As Boolean
    '***************************************************
    'Author: ZaMa
    'Last Modification: 28/04/2010
    'Checks if caster can cast support magic on target user.
    '***************************************************
     
    On Error GoTo Errhandler
 
    With UserList(CasterIndex)
        
        ' Te podes curar a vos mismo
        If CasterIndex = targetIndex Then
            CanSupportUser = True
            Exit Function

        End If
        
        ' No podes ayudar si estas en consulta
        If .flags.EnConsulta Then
            Call WriteConsoleMsg(CasterIndex, "No puedes ayudar usuarios mientras estas en consulta.", FontTypeNames.FONTTYPE_INFO)
            Exit Function

        End If
        
        ' Si estas en la arena, esta todo permitido
        If TriggerZonaPelea(CasterIndex, targetIndex) = TRIGGER6_PERMITE Then
            CanSupportUser = True
            Exit Function

        End If
     
        ' Victima criminal?
        If criminal(targetIndex) Then
        
            ' Casteador Ciuda?
            If Not criminal(CasterIndex) Then
            
                ' Armadas no pueden ayudar
                If esArmada(CasterIndex) Then
                    Call WriteConsoleMsg(CasterIndex, "Los miembros del ejercito real no pueden ayudar a los criminales.", FontTypeNames.FONTTYPE_INFO)
                    Exit Function

                End If
                
                ' Si el ciuda tiene el seguro puesto no puede ayudar
                If .flags.Seguro Then
                    Call WriteConsoleMsg(CasterIndex, "Para ayudar criminales debes sacarte el seguro ya que te volveras criminal como ellos.", FontTypeNames.FONTTYPE_INFO)
                    Exit Function
                Else

                    ' Penalizacion
                    If DoCriminal Then
                        Call VolverCriminal(CasterIndex)
                    Else
                        Call DisNobAuBan(CasterIndex, .Reputacion.NobleRep * 0.5, 10000)

                    End If

                End If

            End If
            
            ' Victima ciuda o army
        Else

            ' Casteador es caos? => No Pueden ayudar ciudas
            If esCaos(CasterIndex) Then
                Call WriteConsoleMsg(CasterIndex, "Los miembros de la legion oscura no pueden ayudar a los ciudadanos.", FontTypeNames.FONTTYPE_INFO)
                Exit Function
                
                ' Casteador ciuda/army?
            ElseIf Not criminal(CasterIndex) Then
                
                ' Esta en estado atacable?
                If UserList(targetIndex).flags.AtacablePor > 0 Then
                    
                    ' No esta atacable por el casteador?
                    If UserList(targetIndex).flags.AtacablePor <> CasterIndex Then
                    
                        ' Si es armada no puede ayudar
                        If esArmada(CasterIndex) Then
                            Call WriteConsoleMsg(CasterIndex, "Los miembros del ejercito real no pueden ayudar a ciudadanos en estado atacable.", FontTypeNames.FONTTYPE_INFO)
                            Exit Function

                        End If
    
                        ' Seguro puesto?
                        If .flags.Seguro Then
                            Call WriteConsoleMsg(CasterIndex, "Para ayudar ciudadanos en estado atacable debes sacarte el seguro, pero te puedes volver criminal.", FontTypeNames.FONTTYPE_INFO)
                            Exit Function
                        Else
                            Call DisNobAuBan(CasterIndex, .Reputacion.NobleRep * 0.5, 10000)

                        End If

                    End If

                End If
    
            End If

        End If

    End With
    
    CanSupportUser = True

    Exit Function
    
Errhandler:
    Call LogError("Error en CanSupportUser, Error: " & Err.Number & " - " & Err.description & " CasterIndex: " & CasterIndex & ", TargetIndex: " & targetIndex)

End Function

Sub UpdateUserHechizos(ByVal UpdateAll As Boolean, _
                       ByVal UserIndex As Integer, _
                       ByVal Slot As Byte)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim LoopC As Byte

    With UserList(UserIndex)

        'Actualiza un solo slot
        If Not UpdateAll Then

            'Actualiza el inventario
            If .Stats.UserHechizos(Slot) > 0 Then
                Call ChangeUserHechizo(UserIndex, Slot, .Stats.UserHechizos(Slot))
            Else
                Call ChangeUserHechizo(UserIndex, Slot, 0)

            End If

        Else

            'Actualiza todos los slots
            For LoopC = 1 To MAXUSERHECHIZOS

                'Actualiza el inventario
                If .Stats.UserHechizos(LoopC) > 0 Then
                    Call ChangeUserHechizo(UserIndex, LoopC, .Stats.UserHechizos(LoopC))
                Else
                    Call ChangeUserHechizo(UserIndex, LoopC, 0)

                End If

            Next LoopC

        End If

    End With

End Sub

Public Function CanSupportNpc(ByVal CasterIndex As Integer, _
                              ByVal targetIndex As Integer) As Boolean
    '***************************************************
    'Author: ZaMa
    'Last Modification: 18/09/2010
    'Checks if caster can cast support magic on target Npc.
    '***************************************************
     
    On Error GoTo Errhandler
 
    Dim OwnerIndex As Integer
 
    With UserList(CasterIndex)
        
        OwnerIndex = Npclist(targetIndex).Owner
        
        ' Si no tiene dueno puede
        If OwnerIndex = 0 Then
            CanSupportNpc = True
            Exit Function

        End If
        
        ' Puede hacerlo si es su propio npc
        If CasterIndex = OwnerIndex Then
            CanSupportNpc = True
            Exit Function

        End If
        
        ' No podes ayudar si estas en consulta
        If .flags.EnConsulta Then
            Call WriteConsoleMsg(CasterIndex, "No puedes ayudar npcs mientras estas en consulta.", FontTypeNames.FONTTYPE_INFO)
            Exit Function

        End If
        
        ' Si estas en la arena, esta todo permitido
        If TriggerZonaPelea(CasterIndex, OwnerIndex) = TRIGGER6_PERMITE Then
            CanSupportNpc = True
            Exit Function

        End If
     
        ' Victima criminal?
        If criminal(OwnerIndex) Then

            ' Victima caos?
            If esCaos(OwnerIndex) Then

                ' Atacante caos?
                If esCaos(CasterIndex) Then
                    ' No podes ayudar a un npc de un caos si sos caos
                    Call WriteConsoleMsg(CasterIndex, "No puedes ayudar npcs que estan luchando contra un miembro de tu faccion.", FontTypeNames.FONTTYPE_INFO)
                    Exit Function

                End If

            End If
        
            ' Uno es caos y el otro no, o la victima es pk, entonces puede ayudar al npc
            CanSupportNpc = True
            Exit Function
                
            ' Victima ciuda
        Else

            ' Atacante ciuda?
            If Not criminal(CasterIndex) Then

                ' Atacante armada?
                If esArmada(CasterIndex) Then

                    ' Victima armada?
                    If esArmada(OwnerIndex) Then
                        ' No podes ayudar a un npc de un armada si sos armada
                        Call WriteConsoleMsg(CasterIndex, "No puedes ayudar npcs que estan luchando contra un miembro de tu faccion.", FontTypeNames.FONTTYPE_INFO)
                        Exit Function

                    End If

                End If
                
                ' Uno es armada y el otro ciuda, o los dos ciudas, puede atacar si no tiene seguro
                If .flags.Seguro Then
                    Call WriteConsoleMsg(CasterIndex, "Para ayudar a criaturas que luchan contra ciudadanos debes sacarte el seguro.", FontTypeNames.FONTTYPE_INFO)
                    Exit Function
                    
                    ' ayudo al npc sin seguro, se convierte en atacable
                Else
                    Call ToogleToAtackable(CasterIndex, OwnerIndex, True)
                    CanSupportNpc = True
                    Exit Function

                End If
                
            End If
            
            ' Atacante criminal y victima ciuda, entonces puede ayudar al npc
            CanSupportNpc = True
            Exit Function
            
        End If
    
    End With
    
    CanSupportNpc = True

    Exit Function
    
Errhandler:
    Call LogError("Error en CanSupportNpc, Error: " & Err.Number & " - " & Err.description & " CasterIndex: " & CasterIndex & ", OwnerIndex: " & OwnerIndex)

End Function

Sub ChangeUserHechizo(ByVal UserIndex As Integer, _
                      ByVal Slot As Byte, _
                      ByVal Hechizo As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    
    UserList(UserIndex).Stats.UserHechizos(Slot) = Hechizo
    
    Call WriteChangeSpellSlot(UserIndex, Slot)

End Sub

Public Sub DesplazarHechizo(ByVal UserIndex As Integer, _
                            ByVal Dire As Integer, _
                            ByVal HechizoDesplazado As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    If (Dire <> 1 And Dire <> -1) Then Exit Sub
    If Not (HechizoDesplazado >= 1 And HechizoDesplazado <= MAXUSERHECHIZOS) Then Exit Sub

    Dim TempHechizo As Integer

    With UserList(UserIndex)

        If Dire = 1 Then 'Mover arriba
            If HechizoDesplazado = 1 Then
                Call WriteConsoleMsg(UserIndex, "No puedes mover el hechizo en esa direccion.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            Else
                TempHechizo = .Stats.UserHechizos(HechizoDesplazado)
                .Stats.UserHechizos(HechizoDesplazado) = .Stats.UserHechizos(HechizoDesplazado - 1)
                .Stats.UserHechizos(HechizoDesplazado - 1) = TempHechizo

            End If

        Else 'mover abajo

            If HechizoDesplazado = MAXUSERHECHIZOS Then
                Call WriteConsoleMsg(UserIndex, "No puedes mover el hechizo en esa direccion.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            Else
                TempHechizo = .Stats.UserHechizos(HechizoDesplazado)
                .Stats.UserHechizos(HechizoDesplazado) = .Stats.UserHechizos(HechizoDesplazado + 1)
                .Stats.UserHechizos(HechizoDesplazado + 1) = TempHechizo

            End If

        End If

    End With

End Sub

Public Sub DisNobAuBan(ByVal UserIndex As Integer, NoblePts As Long, BandidoPts As Long)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    'disminuye la nobleza NoblePts puntos y aumenta el bandido BandidoPts puntos
    Dim EraCriminal As Boolean

    EraCriminal = criminal(UserIndex)
    
    With UserList(UserIndex)

        'Si estamos en la arena no hacemos nada
        If MapData(.Pos.Map, .Pos.X, .Pos.Y).Trigger = eTrigger.ZONAPELEA Then Exit Sub
        
        If .flags.Privilegios And (PlayerType.User Or PlayerType.Consejero) Then
            'pierdo nobleza...
            .Reputacion.NobleRep = .Reputacion.NobleRep - NoblePts

            If .Reputacion.NobleRep < 0 Then
                .Reputacion.NobleRep = 0

            End If
            
            'gano bandido...
            .Reputacion.BandidoRep = .Reputacion.BandidoRep + BandidoPts

            If .Reputacion.BandidoRep > MAXREP Then .Reputacion.BandidoRep = MAXREP
            Call WriteMultiMessage(UserIndex, eMessages.NobilityLost) 'Call WriteNobilityLost(UserIndex)

            If criminal(UserIndex) Then If .Faccion.ArmadaReal = 1 Then Call ExpulsarFaccionReal(UserIndex)

        End If
        
        If Not EraCriminal And criminal(UserIndex) Then
            Call RefreshCharStatus(UserIndex)

        End If

    End With

End Sub
