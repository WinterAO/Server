Attribute VB_Name = "modBanco"
'**************************************************************
' modBanco.bas - Handles the character's bank accounts.
'
' Implemented by Kevin Birmingham (NEB)
' kbneb@hotmail.com
'**************************************************************

'**************************************************************************
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
'**************************************************************************

Option Explicit

Sub IniciarDeposito(ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo Errhandler

    'Hacemos un Update del inventario del usuario
    Call UpdateBanUserInv(True, UserIndex, 0)
    Call WriteBankInit(UserIndex)

    UserList(UserIndex).flags.Comerciando = True

Errhandler:

End Sub

Sub SendBanObj(UserIndex As Integer, Slot As Byte, Object As UserObj)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    UserList(UserIndex).BancoInvent.Object(Slot) = Object

    Call WriteChangeBankSlot(UserIndex, Slot)

End Sub

Sub UpdateBanUserInv(ByVal UpdateAll As Boolean, _
                     ByVal UserIndex As Integer, _
                     ByVal Slot As Byte)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim NullObj As UserObj

    Dim LoopC   As Byte

    With UserList(UserIndex)

        'Actualiza un solo slot
        If Not UpdateAll Then

            'Actualiza el inventario
            If .BancoInvent.Object(Slot).ObjIndex > 0 Then
                Call SendBanObj(UserIndex, Slot, .BancoInvent.Object(Slot))
            Else
                Call SendBanObj(UserIndex, Slot, NullObj)

            End If

        Else

            'Actualiza todos los slots
            For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS

                'Actualiza el inventario
                If .BancoInvent.Object(LoopC).ObjIndex > 0 Then
                    Call SendBanObj(UserIndex, LoopC, .BancoInvent.Object(LoopC))
                Else
                    Call SendBanObj(UserIndex, LoopC, NullObj)

                End If

            Next LoopC

        End If

    End With

End Sub

Sub UserRetiraItem(ByVal UserIndex As Integer, _
                   ByVal BankSlot As Integer, _
                   ByVal Cantidad As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo Errhandler

    Dim ObjIndex As Integer
    Dim InvSlot As Integer

    If Cantidad < 1 Then Exit Sub
    
    Call WriteUpdateUserStats(UserIndex)

    If UserList(UserIndex).BancoInvent.Object(BankSlot).Amount > 0 Then
    
        If Cantidad > UserList(UserIndex).BancoInvent.Object(BankSlot).Amount Then Cantidad = UserList(UserIndex).BancoInvent.Object(BankSlot).Amount
            
        ObjIndex = UserList(UserIndex).BancoInvent.Object(BankSlot).ObjIndex
        
        'Agregamos el obj que compro al inventario
        InvSlot = UserReciveObj(UserIndex, BankSlot, Cantidad)
        
        If InvSlot > 0 Then
            If ObjData(ObjIndex).Log = 1 Then
                Call LogDesarrollo(UserList(UserIndex).Name & " retiro " & Cantidad & " " & ObjData(ObjIndex).Name & "[" & ObjIndex & "]")
            End If
            
            'Actualizamos el inventario del usuario
            Call UpdateUserInv(False, UserIndex, InvSlot)
            'Actualizamos el banco
            Call UpdateBanUserInv(False, UserIndex, BankSlot)
        End If

    End If

Errhandler:

End Sub

Function UserReciveObj(ByVal UserIndex As Integer, _
                  ByVal InvSlot As Integer, _
                  ByVal Cantidad As Integer) As Integer
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim Slot As Integer

    Dim obji As Integer

    With UserList(UserIndex)

        If .BancoInvent.Object(InvSlot).Amount <= 0 Then Exit Function
    
        obji = .BancoInvent.Object(InvSlot).ObjIndex
    
        'Ya tiene un objeto de este tipo?
        Slot = 1

        Do Until .Invent.Object(Slot).ObjIndex = obji And .Invent.Object(Slot).Amount + Cantidad <= MAX_INVENTORY_OBJS
        
            Slot = Slot + 1

            If Slot > .CurrentInventorySlots Then
                Exit Do

            End If

        Loop
    
        'Sino se fija por un slot vacio
        If Slot > .CurrentInventorySlots Then
            Slot = 1

            Do Until .Invent.Object(Slot).ObjIndex = 0
                Slot = Slot + 1

                If Slot > .CurrentInventorySlots Then
                    Call WriteConsoleMsg(UserIndex, "No podes tener mas objetos.", FontTypeNames.FONTTYPE_INFO)
                    Exit Function

                End If

            Loop
            .Invent.NroItems = .Invent.NroItems + 1

        End If
    
        'Mete el obj en el slot
        If .Invent.Object(Slot).Amount + Cantidad <= MAX_INVENTORY_OBJS Then
            'Menor que MAX_INV_OBJS
            .Invent.Object(Slot).ObjIndex = obji
            .Invent.Object(Slot).Amount = .Invent.Object(Slot).Amount + Cantidad
        
            Call QuitarBancoInvItem(UserIndex, InvSlot, Cantidad)

            UserReciveObj = Slot
        Else
            Call WriteConsoleMsg(UserIndex, "No podes tener mas objetos.", FontTypeNames.FONTTYPE_INFO)

        End If

    End With

End Function

Sub QuitarBancoInvItem(ByVal UserIndex As Integer, _
                       ByVal Slot As Byte, _
                       ByVal Cantidad As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim ObjIndex As Integer

    With UserList(UserIndex)
        ObjIndex = .BancoInvent.Object(Slot).ObjIndex

        'Quita un Obj

        .BancoInvent.Object(Slot).Amount = .BancoInvent.Object(Slot).Amount - Cantidad
    
        If .BancoInvent.Object(Slot).Amount <= 0 Then
            .BancoInvent.NroItems = .BancoInvent.NroItems - 1
            .BancoInvent.Object(Slot).ObjIndex = 0
            .BancoInvent.Object(Slot).Amount = 0

        End If

    End With
    
End Sub

Sub UserDepositaItem(ByVal UserIndex As Integer, _
                     ByVal InvSlot As Integer, _
                     ByVal Cantidad As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: 06/04/2020
    '06/04/2020: FrankoH298 - No podemos vender monturas en uso.
    '***************************************************

    Dim ObjIndex As Integer
    Dim BankSlot As Integer
    With UserList(UserIndex)
        If .flags.Equitando = 1 Then
            If .Invent.MonturaEqpSlot = InvSlot Then
                Call WriteConsoleMsg(UserIndex, "No podes depositar tu montura mientras lo estes usando.", FontTypeNames.FONTTYPE_TALK)
                Exit Sub
            End If
        End If
        If .Invent.Object(InvSlot).Amount > 0 And Cantidad > 0 Then
        
            If Cantidad > .Invent.Object(InvSlot).Amount Then Cantidad = .Invent.Object(InvSlot).Amount
            
            ObjIndex = .Invent.Object(InvSlot).ObjIndex
            
            'Agregamos el obj que deposita al banco
            BankSlot = UserDejaObj(UserIndex, InvSlot, Cantidad)
            
            If BankSlot > 0 Then
                If ObjData(ObjIndex).Log = 1 Then
                    Call LogDesarrollo(.Name & " deposito " & Cantidad & " " & ObjData(ObjIndex).Name & "[" & ObjIndex & "]")
                End If
                
                'Actualizamos el inventario del usuario
                Call UpdateUserInv(False, UserIndex, InvSlot)
                
                'Actualizamos el inventario del banco
                Call UpdateBanUserInv(False, UserIndex, BankSlot)
            End If
    
        End If
    End With
End Sub

Function UserDejaObj(ByVal UserIndex As Integer, _
                ByVal InvSlot As Integer, _
                ByVal Cantidad As Integer) As Integer
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim Slot As Integer

    Dim obji As Integer
    
    If Cantidad < 1 Then Exit Function
    
    With UserList(UserIndex)
        obji = .Invent.Object(InvSlot).ObjIndex
        
        'Ya tiene un objeto de este tipo?
        Slot = 1

        Do Until .BancoInvent.Object(Slot).ObjIndex = obji And .BancoInvent.Object(Slot).Amount + Cantidad <= MAX_INVENTORY_OBJS
            Slot = Slot + 1
            
            If Slot > MAX_BANCOINVENTORY_SLOTS Then
                Exit Do

            End If

        Loop
        
        'Sino se fija por un slot vacio antes del slot devuelto
        If Slot > MAX_BANCOINVENTORY_SLOTS Then
            Slot = 1

            Do Until .BancoInvent.Object(Slot).ObjIndex = 0
                Slot = Slot + 1
                
                If Slot > MAX_BANCOINVENTORY_SLOTS Then
                    Call WriteConsoleMsg(UserIndex, "No tienes mas espacio en el banco!!", FontTypeNames.FONTTYPE_INFO)
                    Exit Function

                End If

            Loop
            
            .BancoInvent.NroItems = .BancoInvent.NroItems + 1

        End If
        
        If Slot <= MAX_BANCOINVENTORY_SLOTS Then 'Slot valido

            'Mete el obj en el slot
            If .BancoInvent.Object(Slot).Amount + Cantidad <= MAX_INVENTORY_OBJS Then
                
                'Menor que MAX_INV_OBJS
                .BancoInvent.Object(Slot).ObjIndex = obji
                .BancoInvent.Object(Slot).Amount = .BancoInvent.Object(Slot).Amount + Cantidad
                
                Call QuitarUserInvItem(UserIndex, InvSlot, Cantidad)
                
                UserDejaObj = Slot
            Else
                Call WriteConsoleMsg(UserIndex, "El banco no puede cargar tantos objetos.", FontTypeNames.FONTTYPE_INFO)

            End If

        End If

    End With

End Function

Sub SendUserBovedaTxt(ByVal sendIndex As Integer, ByVal UserIndex As Integer)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error Resume Next

    Dim j As Integer

    Call WriteConsoleMsg(sendIndex, UserList(UserIndex).Name, FontTypeNames.FONTTYPE_INFO)
    Call WriteConsoleMsg(sendIndex, "Tiene " & UserList(UserIndex).BancoInvent.NroItems & " objetos.", FontTypeNames.FONTTYPE_INFO)

    For j = 1 To MAX_BANCOINVENTORY_SLOTS

        If UserList(UserIndex).BancoInvent.Object(j).ObjIndex > 0 Then
            Call WriteConsoleMsg(sendIndex, "Objeto " & j & " " & ObjData(UserList(UserIndex).BancoInvent.Object(j).ObjIndex).Name & " Cantidad:" & UserList(UserIndex).BancoInvent.Object(j).Amount, FontTypeNames.FONTTYPE_INFO)

        End If

    Next

End Sub
