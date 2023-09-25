Attribute VB_Name = "Amigos"
Option Explicit

Public Sub ResetAmigos(ByVal UserIndex As Integer)

    Dim i As Integer

    With UserList(UserIndex)

        For i = 1 To MAXAMIGOS
            .Amigos(i).Nombre = vbNullString
            .Amigos(i).Ignorado = 0
            .Amigos(i).index = 0
        Next i

        .Quien = vbNullString

    End With

End Sub

Public Function NoTieneEspacioAmigos(ByVal UserIndex As Integer) As Boolean

    Dim i     As Long
    Dim Count As Byte

    For i = 1 To MAXAMIGOS

        If LenB(UserList(UserIndex).Amigos(i).Nombre) > 0 Then
            Count = Count + 1
        End If

    Next i

    If Count = MAXAMIGOS Then
        NoTieneEspacioAmigos = True
    End If

End Function

Public Function BuscarSlotAmigoVacio(ByVal UserIndex As Integer) As Byte

    Dim i As Long

    For i = 1 To MAXAMIGOS

        If LenB(UserList(UserIndex).Amigos(i).Nombre) = 0 Then
            BuscarSlotAmigoVacio = i
            Exit Function
        End If

    Next i

End Function

Public Function BuscarSlotAmigoName(ByVal UserIndex As Integer, _
                                    ByVal Nombre As String) As Boolean
    Dim i As Long

    For i = 1 To MAXAMIGOS

        If UCase$(UserList(UserIndex).Amigos(i).Nombre) = UCase$(Nombre) Then
            BuscarSlotAmigoName = True
            Exit Function
        End If

    Next i

End Function

Public Function BuscarSlotAmigoNameSlot(ByVal UserIndex As Integer, _
                                        ByVal Nombre As String) As Byte
    Dim i As Long

    For i = 1 To MAXAMIGOS

        If UCase$(UserList(UserIndex).Amigos(i).Nombre) = UCase$(Nombre) Then
            BuscarSlotAmigoNameSlot = i
            Exit Function
        End If

    Next i

End Function

Public Sub BorrarAmigo(ByVal charName As String, ByVal Amigo As String)
    Dim CharFile As String
    Dim i        As Long
    Dim Tiene    As Boolean
    CharFile = CharPath & charName & ".chr"

    If FileExist(CharFile) Then

        For i = 1 To MAXAMIGOS

            If UCase$(CStr(GetVar(CharFile, "AMIGOS", "NOMBRE" & i))) = UCase$(Amigo) Then
                Tiene = True
                Exit For
            End If

        Next i

        If Tiene Then
            'Lo borramos
            Call WriteVar(CharFile, "AMIGOS", "NOMBRE" & i, vbNullString)
            Call WriteVar(CharFile, "AMIGOS", "IGNORADO" & i, 0)
        End If

    End If

End Sub

Public Function AgregarAmigo(ByVal UserIndex As Integer, _
                                     ByVal Otro As Integer, _
                                     ByRef razon As String) As Boolean

    With UserList(UserIndex)

        If Otro = 0 Or UserIndex = 0 Then
            razon = "Usuario Desconectado"
            AgregarAmigo = False
            Exit Function
        End If

        If UserIndex = Otro Then
            razon = "Usuario Invalido"
            AgregarAmigo = False
            Exit Function
        End If
        
        If EsGm(Otro) = True Then
            razon = "No podes agregar a un Game Master como amigo."
            AgregarAmigo = False
            Exit Function
        End If
        
        If EsGm(UserIndex) = True Then
            razon = "Los Game Masters no pueden agregar a usuarios como amigos."
            AgregarAmigo = False
            Exit Function
        End If
        
        If NoTieneEspacioAmigos(UserIndex) = True Then
            razon = "No tienes mas espacio para poder agregar amigos."
            AgregarAmigo = False
            Exit Function
        End If
        
        If NoTieneEspacioAmigos(Otro) = True Then
            razon = "El otro usuario no tiene mas espacio para aceptar amigos."
            AgregarAmigo = False
            Exit Function
        End If
        
        If BuscarSlotAmigoName(UserIndex, UserList(Otro).Name) = True Then
            razon = "Tu y " & UserList(Otro).Name & " ya son amigos."
            AgregarAmigo = False
            Exit Function

        End If

        AgregarAmigo = True

    End With

End Function

Public Sub ActualizarSlotAmigo(ByVal UserIndex As Integer, _
                               ByVal Slot As Byte, _
                               Optional ByVal Todo As Boolean = False)
    Dim i As Long

    With UserList(UserIndex)

        If Todo Then

            For i = 1 To MAXAMIGOS
                Call WriteCargarListaDeAmigos(UserIndex, i)
            Next i

        Else

            Call WriteCargarListaDeAmigos(UserIndex, Slot)

        End If

    End With

End Sub

Public Function ObtenerIndexLibre(ByVal UserIndex As Integer) As Integer

    Dim i As Long

    For i = 1 To MAXAMIGOS

        If UserList(UserIndex).Amigos(i).index <= 0 Then
            ObtenerIndexLibre = i
            Exit Function
        End If

    Next i

End Function

Public Function ObtenerIndexUsuado(ByVal UserIndex As Integer, _
                                   ByVal Otro As Integer) As Integer
    Dim i As Long

    For i = 1 To MAXAMIGOS

        If UserList(UserIndex).Amigos(i).index = Otro Then
            ObtenerIndexUsuado = i
            Exit Function
        End If

    Next i

End Function

Public Sub ObtenerIndexAmigos(ByVal UserIndex As Integer, ByVal Desconectar As Boolean)
    Dim i    As Long
    Dim Slot As Byte

    With UserList(UserIndex)

        If Desconectar = False Then

            For i = 1 To MAXAMIGOS

                If LenB(UserList(i).Name) > 0 Then

                    If BuscarSlotAmigoName(UserIndex, UserList(i).Name) Then

                        'Lo encontro y agregamos el index
                        Slot = ObtenerIndexLibre(UserIndex)

                        'Por las dudas
                        If Slot > 0 Then .Amigos(Slot).index = i

                        If BuscarSlotAmigoName(i, .Name) Then

                            'Actualizamos la lista del otro
                            Slot = ObtenerIndexLibre(i)

                            If Slot > 0 Then

                                UserList(i).Amigos(Slot).index = UserIndex

                                'Informamos al otro de nuestra presencia
                                Call WriteConsoleMsg(i, "Amigos> " & .Name & " se ha conectado", FontTypeNames.FONTTYPE_CONSEJO)

                            End If

                        End If

                    End If

                End If

            Next i

        Else

            For i = 1 To MAXAMIGOS

                'Antes que nada
                If .Amigos(i).index > 0 Then

                    Call WriteConsoleMsg(.Amigos(i).index, "Amigos> " & .Name & " se ha desconectado", FontTypeNames.FONTTYPE_CONSEJO)

                    'Actualizamos la lista de index de los amigos
                    Slot = ObtenerIndexUsuado(.Amigos(i).index, UserIndex)

                    If Slot > 0 Then UserList(.Amigos(i).index).Amigos(Slot).index = 0

                End If

            Next i

        End If

    End With

End Sub
