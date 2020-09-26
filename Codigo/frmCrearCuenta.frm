VERSION 5.00
Begin VB.Form frmCrearCuenta 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Crear Nueva Cuenta"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
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
   ScaleHeight     =   3015
   ScaleWidth      =   6270
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2520
      Width           =   1575
   End
   Begin VB.CommandButton cmdCrearCuenta 
      BackColor       =   &H0080C0FF&
      Caption         =   "Crear Cuenta"
      Height          =   375
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Frame FraNuevaCuenta 
      BackColor       =   &H00FFC0C0&
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
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6015
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
    
    Salt = RandomString(33)
    
    If SaveNewAccount(TxtNick.Text, txtEmail.Text, oSHA256.SHA256(txtPass.Text & Salt), Salt) Then
        MsgBox "Cuenta " & TxtNick.Text & " creada con exito."
        Unload Me
        
    Else
        MsgBox "Error al crear la cuenta."
        
    End If
    
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub
