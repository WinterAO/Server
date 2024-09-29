VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEventoMapa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Estas variables no se pueden tocar, contienen toda la informacion cargada del evento
Private evNombre As String
Private evMapa As Byte
Private evX As Integer
Private evY As Integer
Private PortalMap As Integer
Private PortalX As Integer
Private PortalY As Integer
Private evDuracion As Long

Private TiempoDuracion As Long
Private TiempoParaEvento As Long

Private EnCurso As Boolean

Public Sub Inicializar(ByVal Nombre As String, _
                       ByVal Mapa As Byte, _
                       ByVal X As Integer, _
                       ByVal Y As Integer, _
                       ByVal pMap As Byte, _
                       ByVal pX As Integer, _
                       ByVal pY As Integer, _
                       ByVal Duracion As Long)

    'Nombre del evento (O nombre del mapa)
    evNombre = Nombre
    'Mapa donde tiene lugar el evento
    evMapa = Mapa
    evX = X
    evY = Y
    'Mapa donde aparece el Portal para ir al evento
    PortalMap = pMap
    PortalX = pX
    PortalY = pY
    'Duracion del evento
    evDuracion = Duracion
    
    Call ResetearTiempos

End Sub

Public Sub IniciarEvento()
    Dim ET As obj
    ET.Amount = 1
    ET.ObjIndex = 378
    
    Call SendData(SendTarget.Toall, 0, PrepareMessageConsoleMsg("[EVENTO] Se ha abierto un portal misterioso en " & MapZonas(PortalMap, MapData(evMapa, evX, evY).ZonaIndex).Name, FontTypeNames.FONTTYPE_SERVER))
    Call SendData(SendTarget.Toall, 0, PrepareMessagePlayWave(SND_EVENTO_PORTAL, NO_3D_SOUND, NO_3D_SOUND))
    Call MakeObj(ET, PortalMap, PortalX, PortalY)
    
    With MapData(PortalMap, PortalX, PortalY)
        .TileExit.Map = evMapa
        .TileExit.X = evX
        .TileExit.Y = evY

    End With
    
    EnCurso = True
End Sub

Public Sub FinalizarEvento()
    Dim X As Integer
    Dim Y As Integer
    
    Call SendData(SendTarget.Toall, 0, PrepareMessageConsoleMsg("[EVENTO] El misterioso portal de " & MapZonas(PortalMap, MapData(evMapa, evX, evY).ZonaIndex).Name & " se ha cerrado.", FontTypeNames.FONTTYPE_SERVER))
    Call EraseObj(10000, evMapa, evX, evY)
    
    EnCurso = False
    
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
        
            If MapData(evMapa, X, Y).UserIndex > 1 Then
                Call WarpUserChar(MapData(evMapa).UserIndex, PortalMap, PortalX + 1, PortalY + 1, True)

                Call WriteConsoleMsg(MapData(evMapa).UserIndex, "Una fuerza misteriosa te expulsa de sus dominios.", FontTypeNames.FONTTYPE_INFO)
            End If
        
        Next X
    Next Y
    
    Call ResetearTiempos
    
End Sub

Public Sub ResetearTiempos()
    TiempoDuracion = evDuracion
End Sub

Public Sub restarTiempo()
    TiempoDuracion = TiempoDuracion - 1
End Sub

Public Property Get getDuracion() As Long
    getDuracion = TiempoDuracion
End Property

Public Property Get getEnCurso() As Boolean
    getEnCurso = EnCurso
End Property

Public Property Get getMapa() As Integer
    getMapa = PortalMap
End Property


