Attribute VB_Name = "Logs"
Option Explicit

Public Sub LogBanFromName(ByVal BannedName As String, _
                   ByVal UserIndex As Integer, _
                   ByVal Motivo As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Call WriteVar(App.Path & "\Dat\" & "BanDetail.dat", BannedName, "BannedBy", UserList(UserIndex).Name)
    Call WriteVar(App.Path & "\Dat\" & "BanDetail.dat", BannedName, "Reason", Motivo)
    
    'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
    Dim mifile As Integer

    mifile = FreeFile
    Open App.Path & "\logs\GenteBanned.log" For Append Shared As #mifile
    Print #mifile, BannedName
    Close #mifile

End Sub

Public Sub LogServerStartTime()

    '*****************************************************************
    'Author: ZaMa
    'Last Modify Date: 15/03/2011
    'Logs Server Start Time.
    '*****************************************************************
    Dim n As Integer

    n = FreeFile
    Open App.Path & "\logs\Main.log" For Append Shared As #n
    Print #n, Date & " " & time & " server iniciado " & GetVersionOfTheServer()
    Close #n

End Sub

Public Sub LogIndex(ByVal index As Integer, ByVal Desc As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo errHandler

    Dim nfile As Integer

    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\" & index & ".log" For Append Shared As #nfile
    Print #nfile, Date & " " & time & " " & Desc
    Close #nfile
    
    Exit Sub

errHandler:

End Sub

Public Sub LogRetos(Desc As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo errHandler

    Dim nfile As Integer

    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\Retos.log" For Append Shared As #nfile
    Print #nfile, Date & " " & time & " " & Desc
    Close #nfile
    
    Exit Sub

errHandler:

End Sub

Public Sub LogStatic(Desc As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo errHandler

    Dim nfile As Integer

    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\Stats.log" For Append Shared As #nfile
    Print #nfile, Date & " " & time & " " & Desc
    Close #nfile

    Exit Sub

errHandler:

End Sub

Public Sub LogTarea(Desc As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo errHandler

    Dim nfile As Integer

    nfile = FreeFile(1) ' obtenemos un canal
    Open App.Path & "\logs\haciendo.log" For Append Shared As #nfile
    Print #nfile, Date & " " & time & " " & Desc
    Close #nfile

    Exit Sub

errHandler:

End Sub

Public Sub LogIP(ByVal str As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim nfile As Integer

    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\IP.log" For Append Shared As #nfile
    Print #nfile, Date & " " & time & " " & str
    Close #nfile

End Sub

Public Sub LogDesarrollo(ByVal str As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    Dim nfile As Integer

    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\desarrollo" & Month(Date) & Year(Date) & ".log" For Append Shared As #nfile
    Print #nfile, Date & " " & time & " " & str
    Close #nfile

End Sub

Public Sub LogAsesinato(texto As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo errHandler

    Dim nfile As Integer
    
    nfile = FreeFile ' obtenemos un canal
    
    Open App.Path & "\logs\asesinatos.log" For Append Shared As #nfile
    Print #nfile, Date & " " & time & " " & texto
    Close #nfile
    
    Exit Sub

errHandler:

End Sub

Public Sub LogHackAttemp(texto As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo errHandler

    Dim nfile As Integer

    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\HackAttemps.log" For Append Shared As #nfile
    Print #nfile, "----------------------------------------------------------"
    Print #nfile, Date & " " & time & " " & texto
    Print #nfile, "----------------------------------------------------------"
    Close #nfile
    
    Exit Sub

errHandler:

End Sub

Public Sub LogCheating(texto As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo errHandler

    Dim nfile As Integer

    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\CH.log" For Append Shared As #nfile
    Print #nfile, Date & " " & time & " " & texto
    Close #nfile
    
    Exit Sub

errHandler:

End Sub

Public Sub LogCriticalHackAttemp(texto As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo errHandler

    Dim nfile As Integer

    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\CriticalHackAttemps.log" For Append Shared As #nfile
    Print #nfile, "----------------------------------------------------------"
    Print #nfile, Date & " " & time & " " & texto
    Print #nfile, "----------------------------------------------------------"
    Close #nfile
    
    Exit Sub

errHandler:

End Sub

Public Sub LogAntiCheat(texto As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo errHandler

    Dim nfile As Integer

    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\AntiCheat.log" For Append Shared As #nfile
    Print #nfile, Date & " " & time & " " & texto
    Print #nfile, ""
    Close #nfile
    
    Exit Sub

errHandler:

End Sub

Public Sub LogCreaciondeCuentas(ByVal Mensaje As String)
    '***************************************************
    'Author: Unknown
    'Last Modification: -
    '
    '***************************************************

    On Error GoTo errHandler

    Dim nfile As Integer

    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\logs\CreaciondeCuentas.log" For Append Shared As #nfile
    Print #nfile, Date & " " & time & Mensaje
    Print #nfile, ""
    Close #nfile
    
    Exit Sub

errHandler:

End Sub
