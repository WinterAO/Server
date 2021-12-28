Attribute VB_Name = "ModInvocaciones"
Option Explicit

Public Const PORTAL_INVOCACION As Integer = 145 'Objeto del portal mientras no sale el bicho
Public Const INVOCACION_INACTIVO_TIME As Long = 180 '3 minutos

Private Type tCoords
    X1 As Integer
    X2 As Integer
    Y1 As Integer
    Y2 As Integer
End Type

Private Type tInvocStatus
    Invocando As Boolean
    Tiempo As Integer
End Type

Private Type tInvocacion
    Mapa As Integer             'Mapa donde se invoca
    NPC As Integer              'Numero del NPC que aparece
    Invocado As Boolean         '¿Esta invocado?
    CastInvocacion As Integer   'Tiempo para la aparicion del titan
    Inactividad As Integer
    Quest As Integer
    Coords(3) As tCoords
    PosAparicion As WorldPos
    EstadoInvocacion As tInvocStatus
    NPCIndex As Integer
End Type

Public Invocacion() As tInvocacion
Public NumInvocaciones As Byte

Public Sub InitInvocaciones()
'***************************************************
'Autor: Lorwik
'Fecha: 19/07/2020
'Descripcion: Cargamos los datos de las invocaciones
'***************************************************

On Error GoTo errHandler

    If frmMain.Visible Then frmMain.txtStatus.Text = "Cargando Invocaciones."
    
    Dim tmpStr As String
    Dim i      As Byte
    Dim j      As Byte
    Dim Leer   As clsIniManager
    Set Leer = New clsIniManager
    
    Call Leer.Initialize(DatPath & "Invocaciones.dat")
    
    NumInvocaciones = val(Leer.GetValue("GLOBAL", "NumInvocaciones"))
    
    frmCargando.cargar.min = 0
    frmCargando.cargar.max = NumInvocaciones
    frmCargando.cargar.Value = 0
    
    ReDim Preserve Invocacion(1 To NumInvocaciones) As tInvocacion
    
    For i = 1 To NumInvocaciones
        
        Invocacion(i).Mapa = val(Leer.GetValue("INVOC" & i, "Map"))
        Invocacion(i).NPC = val(Leer.GetValue("INVOC" & i, "NPC"))
        Invocacion(i).Invocado = False
        Invocacion(i).CastInvocacion = val(Leer.GetValue("INVOC" & i, "CastInvocacion"))
        Invocacion(i).Inactividad = 0
        Invocacion(i).Quest = val(Leer.GetValue("INVOC" & i, "Quest"))
        
        For j = 1 To 3
            
            tmpStr = Leer.GetValue("INVOC" & i, "Coords" & j)
            
            Invocacion(i).Coords(j).X1 = val(ReadField(1, tmpStr, Asc("-")))
            Invocacion(i).Coords(j).X2 = val(ReadField(2, tmpStr, Asc("-")))
            Invocacion(i).Coords(j).Y1 = val(ReadField(3, tmpStr, Asc("-")))
            Invocacion(i).Coords(j).Y2 = val(ReadField(4, tmpStr, Asc("-")))
        
        Next j
        
        tmpStr = Leer.GetValue("INVOC" & i, "PosAparicion")
        
        Invocacion(i).PosAparicion.Map = Invocacion(i).Mapa
        Invocacion(i).PosAparicion.X = val(ReadField(1, tmpStr, Asc("-")))
        Invocacion(i).PosAparicion.Y = val(ReadField(2, tmpStr, Asc("-")))
        
    Next i
    
    Set Leer = Nothing
    
    If frmMain.Visible Then frmMain.txtStatus.Text = Date & " " & time & " - Se inicializaron las invocaciones con exito. Operacion Realizada con exito."
    
    Exit Sub
errHandler:
    MsgBox "error inicializando las invocaciones " & Err.Number & ": " & Err.description
    
End Sub

Public Sub IniciarRitoInvocacion(ByVal UserIndex As Integer)
'***************************************************
'Autor: Lorwik
'Fecha: 19/07/2020
'Descripcion: Iniciamos el rito de invocacion
'***************************************************

    Dim InvocID       As Byte 'Id de la invocacion
    Dim Invocadores(3)  As Integer 'Userindex de los asistente
    Dim X             As Integer
    Dim Y             As Integer
    Dim i             As Byte
    Dim PortalObj     As obj
    
    With UserList(UserIndex)
    
        'Anti corrupcion xD
        If EsGm(UserIndex) Or EsRolesMaster(UserList(UserIndex).Name) Then
            Call WriteConsoleMsg(UserIndex, "Los GM's no pueden invocar.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    
        '¿El usuario que inicia la invocacion esta vivo?
        If .flags.Muerto = 1 Then
            Call WriteMultiMessage(UserIndex, eMessages.UserMuerto)
            Exit Sub

        End If
        
        'Obtenemos el ID de la invocacion
        InvocID = BuscarInvocacion(UserIndex)
        
        '¿El ID es valido?
        If InvocID = 0 Then
            Call WriteConsoleMsg(UserIndex, "No hay ninguna criatura que invocar aquí.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        '¿La invocacion requiere quest?
        If Invocacion(InvocID).Quest = 0 Then
            '¿Tiene la quest en curso?
            If buscarQuestenCurso(UserIndex, Invocacion(InvocID).Quest) > 0 Then
                Call WriteConsoleMsg(UserIndex, "Para invocar esta criatura necesitas tener activa la quest '" & QuestList(Invocacion(InvocID).Quest).Nombre & "'", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        End If
        
        'La criatura ya esta invocada?
        If Invocacion(InvocID).Invocado Then
            Call WriteConsoleMsg(UserIndex, "La criatura ya fue invocada.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        '¿Esta invisible u oculto?
        If .flags.invisible Or .flags.Oculto Then
            Call WriteConsoleMsg(UserIndex, "Para iniciar el ritual debes estar visible.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        For i = 1 To 3
            
            For Y = Invocacion(InvocID).Coords(i).Y1 To Invocacion(InvocID).Coords(i).Y2
                For X = Invocacion(InvocID).Coords(i).X1 To Invocacion(InvocID).Coords(i).X2
                
                    '¿Hay usuario en la posicion?
                    If MapData(Invocacion(InvocID).Mapa, X, Y).UserIndex > 0 Then _
                        MapData(Invocacion(InvocID).Mapa, X, Y).UserIndex = Invocadores(i)
                    
                Next X
            Next Y

            'Si falta uno, cancelamos la invocacion
            If Invocadores(i) = 0 Then
                Call WriteConsoleMsg(UserIndex, "Para iniciar el ritual, 3 personas se deben colocar en las posiciones correctas.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If

        Next i
        
        'Limpiamos el punto de invocacion
        If MapData(Invocacion(InvocID).Mapa, Invocacion(InvocID).PosAparicion.X, Invocacion(InvocID).PosAparicion.Y).ObjInfo.ObjIndex > 0 Then _
            Call EraseObj(MapData(Invocacion(InvocID).PosAparicion.X, Invocacion(InvocID).PosAparicion.Y).ObjInfo.Amount, Invocacion(InvocID).PosAparicion.Map, Invocacion(InvocID).PosAparicion.X, Invocacion(InvocID).PosAparicion.Y)
        
        PortalObj.ObjIndex = PORTAL_INVOCACION
        PortalObj.Amount = 1
        
        'Ponemos la animacion del portal
        Call MakeObj(PortalObj, Invocacion(InvocID).PosAparicion.Map, Invocacion(InvocID).PosAparicion.X, Invocacion(InvocID).PosAparicion.Y)
    
        Invocacion(InvocID).EstadoInvocacion.Invocando = True 'Se esta invocando
        Invocacion(InvocID).EstadoInvocacion.Tiempo = Invocacion(InvocID).CastInvocacion 'Tiempo de casteo para que aparezca el bicho
        Invocacion(InvocID).Invocado = True
        
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Esta dando comienzo el ritual para la invocación de una criatura legendaria.", FontTypeNames.FONTTYPE_FIGHT))
    End With

End Sub

Private Function BuscarInvocacion(ByVal UserIndex As Integer) As Integer
'***************************************************
'Autor: Lorwik
'Fecha: 19/07/2020
'Descripcion: Devuelve la pos del array de la invocacion que se intenta hacer
'***************************************************
    Dim i As Byte
    
    For i = 1 To NumInvocaciones
    
        If UserList(UserIndex).Pos.Map = Invocacion(i).Mapa Then
            BuscarInvocacion = i
            Exit Function
        End If
    
    Next i
    
    'Si llegamos aqui, es que estamos en un mapa invalido
    BuscarInvocacion = 0
    
End Function

Private Function BuscarInvocacionxNPC(ByVal NPCIndex As Integer) As Integer
'***************************************************
'Autor: Lorwik
'Fecha: 19/07/2020
'Descripcion: Devuelve la pos del array de la invocacion que se intenta hacer
'***************************************************
    Dim i As Byte
    
    For i = 1 To NumInvocaciones
    
        If Npclist(NPCIndex).Pos.Map = Invocacion(i).Mapa And Invocacion(i).NPC = NPCIndex Then
            BuscarInvocacionxNPC = i
            Exit Function
        End If
    
    Next i
    
    'Si llegamos aqui, es que estamos en un mapa invalido
    BuscarInvocacionxNPC = 0
    
End Function

Public Sub CastearInvoc(ByVal Indice As Byte)
'***************************************************
'Autor: Lorwik
'Fecha: 19/07/2020
'Descripcion: Casteo de invocaciones en ritual
'***************************************************

    With Invocacion(Indice)

        '¿Hay un NPC invocandose?
        If .EstadoInvocacion.Invocando Then

            '¿Paso el tiempo de casteo?
            If .EstadoInvocacion.Tiempo > 0 Then
                'Si no paso, restamos
                .EstadoInvocacion.Tiempo = .EstadoInvocacion.Tiempo - 1
                
                If .EstadoInvocacion.Tiempo = 10 Then Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("El ritual de invocación casi ha concluido. Faltan 10 segundos para al aparición de la criatura legendaria.", FontTypeNames.FONTTYPE_FIGHT))
                
            Else 'Si paso, invocamos al NPC
                'Eliminamos el portal
                Call EraseObj(MapData(.PosAparicion.Map, .PosAparicion.X, .PosAparicion.Y).ObjInfo.Amount, .PosAparicion.Map, .PosAparicion.X, .PosAparicion.Y)
                
                'Marcamos al NPC como que no se esta invocando
                .EstadoInvocacion.Invocando = False
                
                'El NPC aparece
                Invocacion(Indice).NPCIndex = SpawnNpc(.NPC, .PosAparicion, False, False, False)
                
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Una criatura legendaria ha sido invocada.", FontTypeNames.FONTTYPE_FIGHT))
                
            End If

        End If
        
    End With
End Sub

Public Sub SumarInactividadInvoc(ByVal Indice As Byte)
'***************************************************
'Autor: Lorwik
'Fecha: 19/07/2020
'Descripcion: Sumnamos tiempo de inactividad a los NPC invocados
'***************************************************

    '¿El NPC esta invocado?
    If Invocacion(Indice).Invocado Then
                
        'Sumamos inactividad
        Invocacion(Indice).Inactividad = Invocacion(Indice).Inactividad + 1
                
        '¿El NPC llego al limite de inactividad? Lo borramos
        If Invocacion(Indice).Inactividad >= INVOCACION_INACTIVO_TIME Then _
            Call QuitarNPC(Invocacion(Indice).NPCIndex)
                
    End If
End Sub

Public Sub ResetearInactividadInvoc(ByVal NPCIndex As Integer)
'***************************************************
'Autor: Lorwik
'Fecha: 19/07/2020
'Descripcion: Reseteamos la inactividad de un NPC invocado
'***************************************************

    Dim IdInvoc As Byte
    
    IdInvoc = BuscarInvocacionxNPC(NPCIndex)
    
    Invocacion(IdInvoc).Inactividad = 0

End Sub

Public Sub ResetearInvocacion(ByVal NPCIndex As Integer)
'***************************************************
'Autor: Lorwik
'Fecha: 19/07/2020
'Descripcion: Reseteamos un NPC invocado y que quizas murio
'***************************************************
    
    Dim IdInvoc As Byte
    
    IdInvoc = BuscarInvocacionxNPC(NPCIndex)
    
    With Invocacion(IdInvoc)
        .EstadoInvocacion.Invocando = False 'Se esta invocando
        .EstadoInvocacion.Tiempo = 0 'Tiempo de casteo para que aparezca el bicho
        .Invocado = False
        .Inactividad = 0
        .NPCIndex = 0
    End With
    
End Sub
