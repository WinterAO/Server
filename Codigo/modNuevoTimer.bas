Attribute VB_Name = "modNuevoTimer"
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

'
' Las siguientes funciones devuelven TRUE o FALSE si el intervalo
' permite hacerlo. Si devuelve TRUE, setean automaticamente el
' timer para que no se pueda hacer la accion hasta el nuevo ciclo.
'

Public Function IntervaloPermiteChatGlobal(ByVal UserIndex As Integer, _
                                          Optional ByVal Actualizar As Boolean = True) As Boolean
    '***************************************************
    'Author: Lorwik
    'Last Modification: 25/10/2020
    ' Verificamos si ya puede volver hablar por global
    '***************************************************

    Dim TActual As Long

    TActual = GetTickCount() And &H7FFFFFFF

    With UserList(UserIndex)

        If TActual - .Counters.LastGlobalMsg >= INTERVALO_GLOBAL Then
            If Actualizar Then
                .Counters.LastGlobalMsg = TActual

            End If

            IntervaloPermiteChatGlobal = True
        Else
            IntervaloPermiteChatGlobal = False

        End If

    End With

End Function

Public Function IntervaloNpcVelocidadVariable(ByVal NPCIndex As Integer, _
                                          Optional ByVal Actualizar As Boolean = True) As Boolean
    '***************************************************
    'Author: Lorwik
    'Last Modification: 24/10/2020
    ' Verificamos si la criatura puede comenzar a perseguir
    '***************************************************

    Dim TActual As Long

    TActual = GetTickCount() And &H7FFFFFFF

    With Npclist(NPCIndex)

        If TActual - .Contadores.VelocidadVariable >= .SpeedVar Then
            If Actualizar Then
                .Contadores.VelocidadVariable = TActual

            End If

            IntervaloNpcVelocidadVariable = True
        Else
            IntervaloNpcVelocidadVariable = False

        End If

    End With

End Function

Public Function IntervaloPermiteAtacarNpc(ByVal NPCIndex As Integer, _
                                          Optional ByVal Actualizar As Boolean = True) As Boolean
    '***************************************************
    'Author: Shak
    'Last Modification: 27/07/2016
    ' Verificamos si la criatura puede atacar
    '***************************************************

    Dim TActual As Long

    TActual = GetTickCount() And &H7FFFFFFF

    With Npclist(NPCIndex)

        If TActual - .Contadores.Ataque >= IntervaloNPCPuedeAtacar Then
            If Actualizar Then
                .Contadores.Ataque = TActual

            End If

            IntervaloPermiteAtacarNpc = True
        Else
            IntervaloPermiteAtacarNpc = False

        End If

    End With

End Function

' CASTING DE HECHIZOS
Public Function IntervaloPermiteLanzarSpell(ByVal UserIndex As Integer, _
                                            Optional ByVal Actualizar As Boolean = True) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim TActual As Long

    TActual = GetTickCount() And &H7FFFFFFF

    If TActual - UserList(UserIndex).Counters.TimerLanzarSpell >= IntervaloUserPuedeCastear Then
        If Actualizar Then
            UserList(UserIndex).Counters.TimerLanzarSpell = TActual

        End If

        Call modAntiCheat.RestaCount(UserIndex, 0, 0, 1, 0)
        IntervaloPermiteLanzarSpell = True
    Else
        IntervaloPermiteLanzarSpell = False
        Call modAntiCheat.AddCount(UserIndex, 0, 0, 1, 0)
    End If

End Function

Public Function IntervaloPermiteAtacar(ByVal UserIndex As Integer, _
                                       Optional ByVal Actualizar As Boolean = True) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim TActual As Long

    TActual = GetTickCount() And &H7FFFFFFF

    If TActual - UserList(UserIndex).Counters.TimerPuedeAtacar >= IntervaloUserPuedeAtacar Then
        If Actualizar Then
            UserList(UserIndex).Counters.TimerPuedeAtacar = TActual
            UserList(UserIndex).Counters.TimerGolpeUsar = TActual

        End If

        Call modAntiCheat.RestaCount(UserIndex, 0, 1, 0, 0)
        IntervaloPermiteAtacar = True
    Else
        IntervaloPermiteAtacar = False
        Call modAntiCheat.AddCount(UserIndex, 0, 1, 0, 0)
    End If

End Function

Public Function IntervaloPermiteGolpeUsar(ByVal UserIndex As Integer, _
                                          Optional ByVal Actualizar As Boolean = True) As Boolean
    '***************************************************
    'Author: ZaMa
    'Checks if the time that passed from the last hit is enough for the user to use a potion.
    'Last Modification: 06/04/2009
    '***************************************************

    Dim TActual As Long

    TActual = GetTickCount() And &H7FFFFFFF

    If TActual - UserList(UserIndex).Counters.TimerGolpeUsar >= IntervaloGolpeUsar Then
        If Actualizar Then
            UserList(UserIndex).Counters.TimerGolpeUsar = TActual

        End If

        IntervaloPermiteGolpeUsar = True
    Else
        IntervaloPermiteGolpeUsar = False

    End If

End Function

Public Function IntervaloPermiteMagiaGolpe(ByVal UserIndex As Integer, _
                                           Optional ByVal Actualizar As Boolean = True) As Boolean

    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************
    Dim TActual As Long
    
    With UserList(UserIndex)

        If .Counters.TimerMagiaGolpe > .Counters.TimerLanzarSpell Then
            Exit Function

        End If
        
        TActual = GetTickCount() And &H7FFFFFFF
        
        If TActual - .Counters.TimerLanzarSpell >= IntervaloMagiaGolpe Then
            If Actualizar Then
                .Counters.TimerMagiaGolpe = TActual
                .Counters.TimerPuedeAtacar = TActual
                .Counters.TimerGolpeUsar = TActual

            End If

            IntervaloPermiteMagiaGolpe = True
        Else
            IntervaloPermiteMagiaGolpe = False

        End If

    End With

End Function

Public Function IntervaloPermiteGolpeMagia(ByVal UserIndex As Integer, _
                                           Optional ByVal Actualizar As Boolean = True) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim TActual As Long
    
    If UserList(UserIndex).Counters.TimerGolpeMagia > UserList(UserIndex).Counters.TimerPuedeAtacar Then
        Exit Function

    End If
    
    TActual = GetTickCount() And &H7FFFFFFF
    
    If TActual - UserList(UserIndex).Counters.TimerPuedeAtacar >= IntervaloGolpeMagia Then
        If Actualizar Then
            UserList(UserIndex).Counters.TimerGolpeMagia = TActual
            UserList(UserIndex).Counters.TimerLanzarSpell = TActual

        End If

        IntervaloPermiteGolpeMagia = True
    Else
        IntervaloPermiteGolpeMagia = False

    End If

End Function

' ATAQUE CUERPO A CUERPO
'Public Function IntervaloPermiteAtacar(ByVal UserIndex As Integer, Optional ByVal Actualizar As Boolean = True) As Boolean
'Dim TActual As Long
'
'TActual = GetTickCount() And &H7FFFFFFF''
'
'If TActual - UserList(UserIndex).Counters.TimerPuedeAtacar >= IntervaloUserPuedeAtacar Then
'    If Actualizar Then UserList(UserIndex).Counters.TimerPuedeAtacar = TActual
'    IntervaloPermiteAtacar = True
'Else
'    IntervaloPermiteAtacar = False
'End If
'End Function

' TRABAJO
Public Function IntervaloPermiteTrabajar(ByVal UserIndex As Integer, _
                                         Optional ByVal Actualizar As Boolean = True) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim TActual As Long
    
    TActual = GetTickCount() And &H7FFFFFFF
    
    If TActual - UserList(UserIndex).Counters.TimerPuedeTrabajar >= IntervaloUserPuedeTrabajar Then
        If Actualizar Then UserList(UserIndex).Counters.TimerPuedeTrabajar = TActual
        IntervaloPermiteTrabajar = True
    Else
        IntervaloPermiteTrabajar = False

    End If

End Function

' USAR OBJETOS
Public Function IntervaloPermiteUsar(ByVal UserIndex As Integer, _
                                     Optional ByVal Actualizar As Boolean = True) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: 25/01/2010 (ZaMa)
    '25/01/2010: ZaMa - General adjustments.
    '***************************************************

    Dim TActual As Long
    
    TActual = GetTickCount() And &H7FFFFFFF
    
    If TActual - UserList(UserIndex).Counters.TimerUsar >= IntervaloUserPuedeUsar Then
        If Actualizar Then
            UserList(UserIndex).Counters.TimerUsar = TActual

            'UserList(UserIndex).Counters.failedUsageAttempts = 0
        End If

        Call modAntiCheat.RestaCount(UserIndex, 0, 0, 0, 1)
        IntervaloPermiteUsar = True
    Else
        IntervaloPermiteUsar = False
        Call modAntiCheat.AddCount(UserIndex, 0, 0, 0, 1)
    End If

End Function

Public Function IntervaloPermiteUsarArcos(ByVal UserIndex As Integer, _
                                          Optional ByVal Actualizar As Boolean = True) As Boolean
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim TActual As Long
    
    TActual = GetTickCount() And &H7FFFFFFF
    
    If TActual - UserList(UserIndex).Counters.TimerPuedeUsarArco >= IntervaloFlechasCazadores Then
        If Actualizar Then UserList(UserIndex).Counters.TimerPuedeUsarArco = TActual
        Call modAntiCheat.RestaCount(UserIndex, 1, 0, 0, 0)
        IntervaloPermiteUsarArcos = True
    Else
        IntervaloPermiteUsarArcos = False
        Call modAntiCheat.AddCount(UserIndex, 1, 0, 0, 0)
    End If

End Function

Public Function getInterval(ByVal timeNow As Long, ByVal startTime As Long) As Long ' 0.13.5
    If timeNow < startTime Then
        getInterval = &H7FFFFFFF - startTime + timeNow + 1
    Else
        getInterval = timeNow - startTime
    End If
End Function

Public Function IntervaloPermiteSerAtacado(ByVal UserIndex As Integer, _
                                           Optional ByVal Actualizar As Boolean = False) As Boolean

    '**************************************************************
    'Author: ZaMa
    'Last Modify by: ZaMa
    'Last Modify Date: 13/11/2009
    '13/11/2009: ZaMa - Add the Timer which determines wether the user can be atacked by a NPc or not
    '**************************************************************
    Dim TActual As Long
    
    TActual = GetTickCount() And &H7FFFFFFF
    
    With UserList(UserIndex)

        ' Inicializa el timer
        If Actualizar Then
            .Counters.TimerPuedeSerAtacado = TActual
            .flags.NoPuedeSerAtacado = True
            IntervaloPermiteSerAtacado = False
        Else

            If TActual - .Counters.TimerPuedeSerAtacado >= IntervaloPuedeSerAtacado Then
                .flags.NoPuedeSerAtacado = False
                IntervaloPermiteSerAtacado = True
            Else
                IntervaloPermiteSerAtacado = False

            End If

        End If

    End With

End Function

Public Function IntervaloPerdioNpc(ByVal UserIndex As Integer, _
                                   Optional ByVal Actualizar As Boolean = False) As Boolean

    '**************************************************************
    'Author: ZaMa
    'Last Modify by: ZaMa
    'Last Modify Date: 13/11/2009
    '13/11/2009: ZaMa - Add the Timer which determines wether the user still owns a Npc or not
    '**************************************************************
    Dim TActual As Long
    
    TActual = GetTickCount() And &H7FFFFFFF
    
    With UserList(UserIndex)

        ' Inicializa el timer
        If Actualizar Then
            .Counters.TimerPerteneceNpc = TActual
            IntervaloPerdioNpc = False
        Else

            If TActual - .Counters.TimerPerteneceNpc >= IntervaloOwnedNpc Then
                IntervaloPerdioNpc = True
            Else
                IntervaloPerdioNpc = False

            End If

        End If

    End With

End Function

Public Function IntervaloEstadoAtacable(ByVal UserIndex As Integer, _
                                        Optional ByVal Actualizar As Boolean = False) As Boolean

    '**************************************************************
    'Author: ZaMa
    'Last Modify by: ZaMa
    'Last Modify Date: 13/01/2010
    '13/01/2010: ZaMa - Add the Timer which determines wether the user can be atacked by an user or not
    '**************************************************************
    Dim TActual As Long
    
    TActual = GetTickCount() And &H7FFFFFFF
    
    With UserList(UserIndex)

        ' Inicializa el timer
        If Actualizar Then
            .Counters.TimerEstadoAtacable = TActual
            IntervaloEstadoAtacable = True
        Else

            If TActual - .Counters.TimerEstadoAtacable >= IntervaloAtacable Then
                IntervaloEstadoAtacable = False
            Else
                IntervaloEstadoAtacable = True

            End If

        End If

    End With

End Function

Public Function IntervaloPuedeOcultar(ByVal UserIndex As Integer, _
                                     Optional ByVal Actualizar As Boolean = True) As Boolean
    '**************************************************************
    'Author: Lorwik
    'Last Modify Date: 18/03/2021
    '**************************************************************

    Dim TActual As Long
    
    TActual = GetTickCount() And &H7FFFFFFF
    
    If TActual - UserList(UserIndex).Counters.TimerPuedeOcultar >= IntervaloOcultable Then
        If Actualizar Then
            UserList(UserIndex).Counters.TimerPuedeOcultar = TActual

            'UserList(UserIndex).Counters.failedUsageAttempts = 0
        End If

        Call modAntiCheat.RestaCount(UserIndex, 0, 0, 0, 1)
        IntervaloPuedeOcultar = True
    Else
        IntervaloPuedeOcultar = False
        Call modAntiCheat.AddCount(UserIndex, 0, 0, 0, 1)
    End If

End Function

Public Function checkInterval(ByRef startTime As Long, _
                              ByVal timeNow As Long, _
                              ByVal interval As Long) As Boolean

    Dim lInterval As Long

    If timeNow < startTime Then
        lInterval = &H7FFFFFFF - startTime + timeNow + 1
    Else
        lInterval = timeNow - startTime

    End If

    If lInterval >= interval Then
        startTime = timeNow
        checkInterval = True
    Else
        checkInterval = False

    End If

End Function

