Attribute VB_Name = "modFortalezas"
Option Explicit

'Numero maximo de estados que puede tener una puerta
Public Const ESTADOSPUERTA As Byte = 3

Public NumFortalezas As Byte

Public Fortaleza() As clsFortalezas

Public Function CargarFortalezas() As Boolean
    '************************************************
    'Autor: Lorwik
    'Fecha: 07/11/2020
    'Descripcion: Carga las propiedades de cada fortaleza y las manda a inicializar
    '************************************************
    
    Dim NombreFortaleza                     As String
    Dim MapaFortaleza                       As Integer
    Dim ReyIndex                            As Integer
    Dim XRey                                As Integer
    Dim YRey                                As Integer
    Dim PuertaIndex                         As Integer
    Dim XPuerta                             As Integer
    Dim YPuerta                             As Integer
    Dim EstadoPuerta(1 To ESTADOSPUERTA)    As Integer
    Dim ClanConquistador                    As String
    Dim Recompensa                          As Long
    Dim Intervalo                           As Long
    
    Dim LoopC                               As Byte
    Dim i                                   As Byte
    
    Dim Leer                                As clsIniManager
    
    Set Leer = New clsIniManager
    
    Call Leer.Initialize(DatPath & "Fortalezas.dat")
    
    NumFortalezas = val(Leer.GetValue("INIT", "NumFortalezas"))
    
    If NumFortalezas = 0 Then
        If frmMain.Visible Then frmMain.txtStatus.Text = "No hay fortalezas."
        Exit Function
    End If
    
    'Redimensionamos el array
    ReDim Fortaleza(1 To NumFortalezas) As clsFortalezas
    
    For LoopC = 1 To NumFortalezas
    
        NombreFortaleza = Leer.GetValue("FORTALEZA" & LoopC, "Nombre")
        MapaFortaleza = val(Leer.GetValue("FORTALEZA" & LoopC, "Mapa"))
        ReyIndex = val(Leer.GetValue("FORTALEZA" & LoopC, "ReyIndex"))
        XRey = val(Leer.GetValue("FORTALEZA" & LoopC, "XRey"))
        YRey = val(Leer.GetValue("FORTALEZA" & LoopC, "YRey"))
        PuertaIndex = val(Leer.GetValue("FORTALEZA" & LoopC, "PuertaIndex"))
        XPuerta = val(Leer.GetValue("FORTALEZA" & LoopC, "XPuerta"))
        YPuerta = val(Leer.GetValue("FORTALEZA" & LoopC, "YPuerta"))
        
        For i = 1 To ESTADOSPUERTA
            EstadoPuerta(i) = val(Leer.GetValue("FORTALEZA" & LoopC, "EstadoPuerta" & i))
        Next i
        
        ClanConquistador = Leer.GetValue("FORTALEZA" & LoopC, "ClanConquistador")
        Recompensa = val(Leer.GetValue("FORTALEZA" & LoopC, "Recompensa"))
        Intervalo = val(Leer.GetValue("FORTALEZA" & LoopC, "IntervaloRecompensa"))
    
        Set Fortaleza(LoopC) = New clsFortalezas
        
        Call Fortaleza(LoopC).Inicializar(NombreFortaleza, MapaFortaleza, _
                        ReyIndex, XRey, YRey, _
                        PuertaIndex, XPuerta, YPuerta, EstadoPuerta(), _
                        ClanConquistador, Recompensa, Intervalo)
    
    Next LoopC
    
    Set Leer = Nothing

    If frmMain.Visible Then frmMain.txtStatus.Text = Date & " " & time & " - Fortalezas cargadas con exito."
    
    CargarFortalezas = True
    Exit Function

errHandler:
    MsgBox "Error cargando fortalezas.dat " & Err.Number & ": " & Err.description
    CargarFortalezas = False
    
End Function

Public Sub DestruirFortalezas()
    '************************************************
    'Autor: Lorwik
    'Fecha: 07/11/2020
    'Descripcion: Destruye todas las fortalezas cargadas
    '************************************************

    Dim LoopC As Byte
    
    For LoopC = 1 To NumFortalezas
        Set Fortaleza(LoopC) = Nothing
    Next LoopC
    
    NumFortalezas = 0
    
End Sub

Public Function PuedeAtacarFortaleza(ByVal UserIndex As Integer, ByVal NFortaleza As Byte) As Boolean
    '************************************************
    'Autor: Lorwik
    'Fecha: 07/11/2020
    'Descripcion: Devuelve True o False si puede atacar una fortaleza
    '************************************************
    
    With UserList(UserIndex)
    
        '¿Tiene clan?
        If UserList(UserIndex).GuildIndex = 0 Then
            Call WriteConsoleMsg(UserIndex, "Para atacar una fortaleza necesitas pertenece a un clan.", FontTypeNames.FONTTYPE_INFO)
            PuedeAtacarFortaleza = False
            Exit Function
        End If
    
        '¿Su clan es el dueño?
        If modGuilds.GuildName(.GuildIndex) = Fortaleza(NFortaleza).ClanConquistador Then
            Call WriteConsoleMsg(UserIndex, "Tu clan ya controla esta fortaleza.", FontTypeNames.FONTTYPE_INFO)
            PuedeAtacarFortaleza = False
            Exit Function
        End If
        
        PuedeAtacarFortaleza = True
    
    End With
    
End Function

Public Function IndiceFortaleza(ByVal UserIndex As Integer) As Integer
    '************************************************
    'Autor: Lorwik
    'Fecha: 07/11/2020
    'Descripcion: Devuelve el indice de la fortaleza en la que se encuentra actualmente.
    'si devuelve 0 es que no esta en ninguna fortaleza
    '************************************************
    
    Dim i As Byte
    
    With UserList(UserIndex)
        
        For i = 1 To NumFortalezas
            'El mapa en el que se encuentra coincide con el de alguna fortaleza?
            If .Pos.Map = Fortaleza(i).MapaFortaleza Then
                IndiceFortaleza = i
                Exit Function
            End If
        Next i
    
        'Si llegamos aqui es que no esta en ninguna fortaleza
        IndiceFortaleza = 0
    
    End With
End Function
