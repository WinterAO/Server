VERSION 5.00
Begin VB.Form frmCrearCuenta 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Crear Nueva Cuenta"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6270
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   6270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraEditarPass 
      BackColor       =   &H00000000&
      Caption         =   "Editar Pass Cuenta"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   2085
      Left            =   90
      TabIndex        =   9
      Top             =   2760
      Width           =   6015
      Begin VB.TextBox txtAccountName 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1920
         TabIndex        =   12
         Top             =   480
         Width           =   3495
      End
      Begin VB.TextBox txtNewPass 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1920
         TabIndex        =   11
         Top             =   960
         Width           =   3495
      End
      Begin VB.CommandButton cmdEditarPass 
         BackColor       =   &H008080FF&
         Caption         =   "Cambiar Pass Cuenta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   540
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1560
         Width           =   4935
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre de Cuenta: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   600
         Width           =   1650
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nueva Contraseña:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   1020
         Width           =   1590
      End
   End
   Begin VB.CommandButton cmdCancelar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3990
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5010
      Width           =   2085
   End
   Begin VB.Frame FraNuevaCuenta 
      BackColor       =   &H00000000&
      Caption         =   "Nueva Cuenta"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6015
      Begin VB.CommandButton cmdCrearCuenta 
         BackColor       =   &H0080C0FF&
         Caption         =   "Crear Cuenta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   540
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2040
         Width           =   4935
      End
      Begin VB.TextBox txtPass 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1920
         TabIndex        =   6
         Top             =   1440
         Width           =   3495
      End
      Begin VB.TextBox txtEmail 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1920
         TabIndex        =   5
         Top             =   960
         Width           =   3495
      End
      Begin VB.TextBox TxtNick 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1920
         TabIndex        =   1
         Top             =   480
         Width           =   3495
      End
      Begin VB.Label lblContraseña 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contraseña:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   195
         Left            =   840
         TabIndex        =   4
         Top             =   1560
         Width           =   1020
      End
      Begin VB.Label lblEmail 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Email:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   195
         Left            =   1320
         TabIndex        =   3
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label lblNombreDe 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre de Cuenta: "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   1650
      End
   End
End
Attribute VB_Name = "frmCrearCuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCrearCuenta_Click()
    
    Dim Salt As String
    
    Dim oSHA256 As CSHA256

    Set oSHA256 = New CSHA256
    
    If LenB(TxtNick.Text) > 24 Or LenB(TxtNick.Text) = 0 Then
        MsgBox "Nombre invalido."
        Exit Sub

    End If
    
    If LenB(txtEmail.Text) = 0 Then
        MsgBox "Escribe un Email"
        Exit Sub
    End If
    
    If LenB(txtPass.Text) = 0 Then
        MsgBox "Escribe una contraseña"
        Exit Sub
    End If
    
    If CuentaExisteDatabase(TxtNick.Text) Then
        MsgBox "El nombre de la cuenta ya existe"
        Exit Sub
    End If
    
    Salt = RandomString(32)
    
    If SaveNewAccount(TxtNick.Text, txtEmail.Text, oSHA256.SHA256(txtPass.Text & Salt), Salt) Then
        MsgBox "Cuenta " & TxtNick.Text & " creada con exito."
        
    Else
        MsgBox "Error al crear la cuenta."
    
    End If
    
End Sub

Private Sub cmdEditarPass_Click()
    Dim Salt As String
    
    Dim oSHA256 As CSHA256

    Set oSHA256 = New CSHA256
    
    If LenB(txtAccountName.Text) > 24 Or LenB(txtAccountName.Text) = 0 Then
        MsgBox "Nombre invalido."
        Exit Sub

    End If
    
    If LenB(txtAccountName.Text) = 0 Then
        MsgBox "Escribe un nombre de cuenta"
        Exit Sub
    End If
    
    If LenB(txtNewPass.Text) = 0 Then
        MsgBox "Escribe una contraseña"
        Exit Sub
    End If
    
    If Not CuentaExisteDatabase(txtAccountName.Text) Then
        MsgBox "El nombre de la cuenta no existe"
        Exit Sub
    End If
    
    Salt = RandomString(32)
    
    If SaveAccountEditPassDatabase(txtAccountName.Text, oSHA256.SHA256(txtNewPass.Text & Salt), Salt) Then
        MsgBox "Password de la cuenta " & txtAccountName.Text & " cambiada con exito."
        
    Else
        MsgBox "Error al cambiar password de la cuenta " & txtAccountName.Text & "."
    
    End If
    
End Sub


Private Sub cmdCancelar_Click()
    Unload Me
End Sub

