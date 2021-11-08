VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDataBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Database_DataSource As String
Private Database_Host       As String
Private Database_Name       As String
Private Database_Username   As String
Private Database_Password   As String

Public Database_Connection  As ADODB.Connection
Public Database_RecordSet   As ADODB.Recordset
Public Database_Command     As ADODB.Command

'CONSTRUCTOR
Public Sub Inicialiar(ByVal source As String, _
                      ByVal Host As String, _
                      ByVal Name As String, _
                      ByVal User As String, _
                      ByVal Pass As String)
                      
    
    Database_DataSource = source
    Database_Host = Host
    Database_Name = Name
    Database_Username = User
    Database_Password = Pass

End Sub

Public Function Database_Connect() As Boolean
    '************************************************************************************
    'Author: Juan Andres Dalmasso
    'Last Modification: 21/09/2019
    '21/09/2019 Jopi - Agregue soporte a conexion via DSN. Solo para usuarios avanzados.
    '************************************************************************************
    
    On Error GoTo ErrorHandler
    
    Set Database_Connection = New ADODB.Connection
    
    If Len(Database_DataSource) <> 0 Then
    
        Database_Connection.ConnectionString = "DATA SOURCE=" & Database_DataSource & ";"
        
    Else
    
        Database_Connection.ConnectionString = "DRIVER={MySQL ODBC 8.0 ANSI Driver};" & _
                                               "SERVER=" & Database_Host & ";" & _
                                               "DATABASE=" & Database_Name & ";" & _
                                               "USER=" & Database_Username & ";" & _
                                               "PASSWORD=" & Database_Password & ";" & _
                                               "OPTION=3"
    End If
    
    'Debug.Print Database_Connection.ConnectionString
    
    Database_Connection.CursorLocation = adUseClient
    Database_Connection.Open
    
    Database_Connect = True
    
    Exit Function
    
ErrorHandler:
    Call LogDatabaseError("Database Error: " & Err.Number & " - " & Err.description)
    Debug.Print "Database Error: " & Err.Number & " - " & Err.description
    Database_Connect = False

End Function

Public Sub Database_Close()

    '***************************************************
    'Author: Juan Andres Dalmasso
    'Last Modification: 18/09/2018
    '***************************************************
    On Error GoTo ErrorHandler
     
    Database_Connection.Close
    Set Database_Connection = Nothing
     
    Exit Sub
     
ErrorHandler:
    Call LogDatabaseError("Unable to close Mysql Database: " & Err.Number & " - " & Err.description)

End Sub

Public Function Database_Reconnect() As Boolean
'***************************************************
'Author: Lorwik
'Fecha: 13/09/2020
'Descripcion: Reconexión de la base de datos
'***************************************************

    Dim CerrolaConexion As Boolean
    
    'Si la conexión ya existia...
    If Database_Connection Is Nothing Then
        CerrolaConexion = False
        
    Else
        Call Database_Close
        CerrolaConexion = True
        
    End If
    
    Call Database_Connect
    
    Call LogDatabaseError("Base de datos reconectada. ¿Se cerro la conexion?: " & CerrolaConexion)
    
End Function

Public Function CheckSQLStatus() As Boolean
    '***************************************************
    'Author: Lorwik
    'Fecha: 17/07/2020
    'Descripcion: Comprobamos el estado de la conexion a la base de datos.
    '***************************************************
    
    If Database_Connection Is Nothing Then
        CheckSQLStatus = False
        Exit Function
    End If
    
    If Database_Connection.State = 0 Then
        CheckSQLStatus = False
        Exit Function
    End If
    
    'La conexion es correcta
    CheckSQLStatus = True

End Function

Public Function MakeQuery(query As String, ByVal NoResult As Boolean, ParamArray Query_Parameters() As Variant) As Boolean
    ' 17/10/2020 Autor: Alexis Caraballo (WyroX)
    ' Hace una unica query a la db. Asume una conexion.
    ' Si NoResult = False, el metodo lee el resultado de la query
    ' Guarda el resultado en QueryData
    
    On Error GoTo ErrorHandler
    
    Dim query_Params As Variant
    Dim RecordsAffected As Long
    
    Set Database_Command = New ADODB.Command

    If UBound(Query_Parameters) < 0 Then
        query_Params = Null
    Else
        query_Params = Query_Parameters

    End If

    With Database_Command
    
        ' Clear old params
        Dim i As Integer
        For i = 0 To .Parameters.Count - 1
            Call .Parameters.Delete(0)
        Next
        
        .ActiveConnection = Database_Connection
        .CommandText = query

        If NoResult Then
            Call .Execute(RecordsAffected, query_Params, adExecuteNoRecords)
    
        Else
            Set Database_RecordSet = .Execute(RecordsAffected, query_Params)
    
            If Database_RecordSet.BOF Or Database_RecordSet.EOF Then
                Set Database_RecordSet = Nothing
                Set Database_Command = Nothing
                MakeQuery = False
                Exit Function
            End If
    
        End If
        
    End With
    
    Set Database_Command = Nothing
    MakeQuery = True
    Exit Function
    
ErrorHandler:

    MakeQuery = False

End Function

