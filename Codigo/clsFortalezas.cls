VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFortalezas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Nombre que tendra la fortaleza
Private Nombre_Fortaleza As String

'Mapa en el que se encuenta la fortaleza
Private Mapa_Fortaleza As Integer

'Numero en el NPC.dat del rey o defensor
Private Rey As Integer

'NPCIndex del rey
Private ReyIndex As Integer

'Posicion del rey en el mapa
Private Rey_Pos As WorldPos

'Numero que tendra en el NPC.dat la pueta
Private Puerta As Integer

'NPCIndex de la puerta
Private PuertaIndex As Integer

'Posicion de la puerta en el mapa
Private Puerta_Pos As WorldPos

'Numero de los cuerpos en los que ira cambiando segun la vida de la puerta.
Private Estado_Puerta(1 To ESTADOSPUERTA) As Integer

'La puerta fue derribada o no
Private Puerta_Destruida As Boolean

'Nombre del clan dueño de la fortaleza
Private Clan_Conquistador As String

'Cantida de oro que obtendra el clan conquistador
Private Oro_Recompensa As Long

'Intervalo en el que se le ira dando el orog al clan conquistador
Private Intervalo_Recompensa As Long

'Indica si el rey murio
Private Conquistado As Boolean

'METEODOS GET
Public Property Get NombreFortaleza() As String
    NombreFortaleza = Nombre_Fortaleza
End Property

Public Property Get MapaFortaleza() As String
    MapaFortaleza = Mapa_Fortaleza
End Property

Public Property Get ReyFortaleza() As Integer
    ReyFortaleza = Rey
End Property

Public Property Get PuertaFortaleza() As Integer
    PuertaFortaleza = Puerta
End Property

Public Property Get ClanConquistador() As String
    ClanConquistador = Clan_Conquistador
End Property

'CONSTRUCTOR
Public Sub Inicializar(ByVal NombreFortaleza As String, _
                        ByVal MapaFortaleza As Integer, _
                        ByVal NRey As Integer, _
                        ByVal XRey As Integer, _
                        ByVal YRey As Integer, _
                        ByVal NPuerta As Integer, _
                        ByVal XPuerta As Integer, _
                        ByVal YPuerta As Integer, _
                        ByRef EstadoPuerta() As Integer, _
                        ByVal ClanConquistador As String, _
                        ByVal Recompensa As Long, _
                        ByVal Intervalo As Long)
                        
    '**************************************************
    'Autor: Lorwik
    'Fecha: 07/11/2020
    'Descripción: Inicializa una nueva fortaleza
    '**************************************************
                        
    Dim i As Byte
                        
    Nombre_Fortaleza = NombreFortaleza
    Mapa_Fortaleza = MapaFortaleza
    
    Rey = NRey
    Rey_Pos.Map = MapaFortaleza
    Rey_Pos.X = XRey
    Rey_Pos.Y = YRey

    Puerta = NPuerta
    Puerta_Pos.Map = MapaFortaleza
    Puerta_Pos.X = XPuerta
    Puerta_Pos.Y = YPuerta
    
    For i = 1 To ESTADOSPUERTA
        Estado_Puerta(i) = EstadoPuerta(i)
    Next i
    
    Clan_Conquistador = ClanConquistador
    Oro_Recompensa = Recompensa
    Intervalo_Recompensa = Intervalo

    'Iniciamos los estados globales
    Conquistado = False
    Puerta_Destruida = False

    'Spawn Rey
    ReyIndex = SpawnNpc(Rey, Rey_Pos, True, False)
    'Spawn Puerta
    PuertaIndex = SpawnNpc(Puerta, Puerta_Pos, True, False)

End Sub

Public Sub ReiniciarCastillo()
    '**************************************************
    'Autor: Lorwik
    'Fecha: 07/11/2020
    'Descripción: Reinicia la fortaleza
    '**************************************************
    
    'Si el Castillo no fue conquistado, es por que el rey sigue vivo, lo quitamos
    If Conquistado = False Then
        Call QuitarNPC(ReyIndex)
    Else
        Conquistado = False
    End If
    
    'Si la puerta no fue destruida la quitamos
    If Puerta_Destruida = False Then
        Call QuitarNPC(PuertaIndex)
    Else
        Puerta_Destruida = False
    End If
    
    'Spawn Rey
    ReyIndex = SpawnNpc(Rey, Rey_Pos, True, False)
    'Spawn Puerta
    PuertaIndex = SpawnNpc(Puerta, Puerta_Pos, True, False)
    
End Sub
