VERSION 5.00
Begin VB.Form frmUserList 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Debug de Userlist"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   5520
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDesconectar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Desconectar Cuenta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4680
      Width           =   2175
   End
   Begin VB.CommandButton cmdDesconectar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Desconectar PJ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4680
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Echar todos los no Logged"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4200
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2400
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   3000
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   2400
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   120
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Actualiza"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4200
      Width           =   2175
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3960
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmUserList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************
' frmUserList.frm
'
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

Private Sub cmdDesconectar_Click(index As Integer)
'********************************************
'Autor: Lorwik
'Fecha: 13/07/2020
'Descripción: Forzamos la desconexion de un PJ o cuenta del server
'********************************************

    Dim UserIndex As Integer
    
    UserIndex = List1.ItemData(List1.ListIndex)
    
    If UserIndex <= 0 Then
        MsgBox "Usuario invalido"
        Exit Sub
    End If
    
    Select Case index
    
        Case 0 'Desconectar PJ
            If MsgBox("¿Seguro que quieres forzar la desconexion de este PJ?", vbYesNo, "¡DESCONEXION PJ!") = vbNo Then Exit Sub
            
            '¿Tiene un PJ Conectado?
            If UserList(UserIndex).flags.UserLogged Then
                Call WriteErrorMsg(UserIndex, "Has sido desconectado del servidor.")
                Call WriteDisconnect(UserIndex)
                Call FlushBuffer(UserIndex)
                Call CloseUser(UserIndex)
            
            Else
                MsgBox "El Usuario no tiene ningun PJ conectado"
                
            End If
                        
        Case 1 'Desconectar cuenta
            If MsgBox("¿Seguro que quieres forzar la desconexion de esta Cuenta? Si el PJ esta conectado se le expulsara", vbYesNo, "¡DESCONEXION CUENTA!") = vbNo Then Exit Sub
            
            '¿Tiene la cuenta conectada?
            If UserList(UserIndex).flags.AccountLogged Then
                Call WriteErrorMsg(UserIndex, "Tu cuenta ha sido desconectada del servidor. Prueba a relogear.")
                
                Call WriteDisconnect(UserIndex)
                Call FlushBuffer(UserIndex)
                
                '¿Tiene un PJ Conectado?
                If UserList(UserIndex).flags.UserLogged Then _
                    Call CloseUser(UserIndex)
                    
                Call CloseAccount(UserIndex)
                
            Else
                MsgBox "El Usuario no tiene ninguna cuenta conectado"
                
            End If
                
    End Select
    
End Sub

Private Sub Command1_Click()

    Dim LoopC As Integer

    Text2.Text = "MaxUsers: " & MaxUsers & vbCrLf
    Text2.Text = Text2.Text & "LastUser: " & LastUser & vbCrLf
    Text2.Text = Text2.Text & "NumUsers: " & NumUsers & vbCrLf
    'Text2.Text = Text2.Text & "" & vbCrLf

    List1.Clear

    For LoopC = 1 To MaxUsers
        List1.AddItem Format(LoopC, "000") & " " & IIf(UserList(LoopC).flags.UserLogged, UserList(LoopC).Name, "")
        List1.ItemData(List1.NewIndex) = LoopC
    Next LoopC

End Sub

Private Sub Command2_Click()

    Dim LoopC As Integer

    For LoopC = 1 To MaxUsers

        If UserList(LoopC).ConnID <> -1 And Not UserList(LoopC).flags.UserLogged Then
            Call CloseUser(LoopC)

        End If

    Next LoopC

End Sub

Private Sub List1_Click()

    Dim UserIndex As Integer

    If List1.ListIndex <> -1 Then
        UserIndex = List1.ItemData(List1.ListIndex)

        If UserIndex > 0 And UserIndex <= MaxUsers Then

            With UserList(UserIndex)
                Text1.Text = "UserLogged: " & .flags.UserLogged & vbCrLf
                Text1.Text = Text1.Text & "AccountLogged: " & .flags.AccountLogged & vbCrLf
                Text1.Text = Text1.Text & "IdleCount: " & .Counters.IdleCount & vbCrLf
                Text1.Text = Text1.Text & "ConnId: " & .ConnID & vbCrLf
                Text1.Text = Text1.Text & "ConnIDValida: " & .ConnIDValida & vbCrLf

            End With

        End If

    End If

End Sub
