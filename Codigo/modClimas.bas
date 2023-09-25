Attribute VB_Name = "modClimas"
'********************************Modulo Climas*********************************
'Autor: Lorwik
'Last Modification: 09/082020
'Controla el clima y lo envia al cliente.
'Nota: Cuando reformemos el sistema de lluvias, todo va a ir aqui.
'******************************************************************************
Option Explicit

Enum eColorEstado
    Amanecer = 0
    MedioDia
    Tarde
    Noche
    Lluvia
    Niebla
    FogLluvia  'Niebla mas lluvia
End Enum

Public DayStatus As eColorEstado 'Establece el color actual del dia

'Todo en minutos:
'Porcentaje del 1 al 100 de la lluvia sea con niebla
Private Const FogProb As Byte = 20 'Niebla
Private Const FogLluviaProb As Byte = 5 'Nieva + Lluvia

Public Sub SortearHorario(Optional ByVal Clima As eColorEstado)
'***************************************************************************************
'Autor: Lorwik
'Ultima modificación: 23/12/2018
'Descripción: Sorteamos el clima, si hay tormenta y es de Mañana o de Dia
'ponemos el efecto de tarde, pero si es de Tarde o de Noche no ponemos nigun efecto.
'***************************************************************************************

    'Si esta lloviendo ignoramos el resto y solo mandamos el estado lluvia
    If Lloviendo Then
    
        'Solo para el Main del server:
        Select Case Clima
        
            Case eColorEstado.Lluvia
                frmMain.lblLloviendoInfo.Caption = "Hora: Lloviendo - [" & Hour(Now) & ":" & Minute(Now) & "]"
                
            Case eColorEstado.Niebla
                frmMain.lblLloviendoInfo.Caption = "Hora: Niebla - [" & Hour(Now) & ":" & Minute(Now) & "]"
                
            Case eColorEstado.FogLluvia
                frmMain.lblLloviendoInfo.Caption = "Hora: Niebla + Lluvia - [" & Hour(Now) & ":" & Minute(Now) & "]"
                
        End Select
        
        Call ColorClima(Clima)
        
    Else
        If (Hour(Now) >= 5 And Hour(Now) < 10) Then 'Amanecer
            Call ColorClima(eColorEstado.Amanecer)
            frmMain.lblLloviendoInfo.Caption = "Hora: Mañana - [" & Hour(Now) & ":" & Minute(Now) & "]"
            
        ElseIf (Hour(Now) >= 11 And Hour(Now) < 16) Then  'MedioDia
            Call ColorClima(eColorEstado.MedioDia)
            frmMain.lblLloviendoInfo.Caption = "Hora: MedioDia - [" & Hour(Now) & ":" & Minute(Now) & "]"
            
        ElseIf (Hour(Now) >= 17 And Hour(Now) < 23) Then 'Tarde
            Call ColorClima(eColorEstado.Tarde)
            frmMain.lblLloviendoInfo.Caption = "Hora: Tarde - [" & Hour(Now) & ":" & Minute(Now) & "]"
            
        ElseIf (Hour(Now) >= 0 And Hour(Now) < 4) Then 'Noche
            Call ColorClima(eColorEstado.Noche)
            frmMain.lblLloviendoInfo.Caption = "Hora: Noche - [" & Hour(Now) & ":" & Minute(Now) & "]"
        End If
        
    End If
    
End Sub

'Enviamos el Clima
Private Sub ColorClima(Clima As eColorEstado)
'****************************************
'Autor: Lorwik
'Ultima modificación: 09/08/2020
'Enviamos el clima
'****************************************

    Dim UserIndex As Integer
    Dim i As Long
    
    DayStatus = Clima

    Call SendData(SendTarget.ToAll, 0, PrepareMessageActualizarClima())
    
End Sub

Public Sub SortearClima(Optional ByVal Forzar As Byte = 0)
'**********************************************
'Autor: Lorwik
'Ultima modificación: 09/08/2020
'Descripción: En este sub vamos a sortear si va lloviar, va hacer niebla o ambas cosas
'**********************************************

    Dim Clima As eColorEstado
    Dim DadosAleatorios As Integer
    
    '¿Esta lloviendo?
    If Lloviendo Then

        If Forzar = 0 Then
            'Por el momento seteamos la lluvia, ya que no requiere probs
            Clima = eColorEstado.Lluvia
            
            'Vamos a tirar los datos
            DadosAleatorios = RandomNumber(1, 1000)
            
            '¿Va haber niebla?
            If FogProb >= DadosAleatorios Then
            
                'Ok, seteamos niebla
                Clima = eColorEstado.Niebla
            
                'Este porcentaje siempre es menor ¿lo pasara?
                If FogLluviaProb >= DadosAleatorios Then
                    '¡Premio! Se vieneeee....
                    Clima = eColorEstado.FogLluvia
                End If
                
            End If
            
        Else '¿Queremos forzar la aparicion de algun fenomeno?
        
            Select Case Forzar
            
                Case 1
                    Clima = eColorEstado.Lluvia
                    
                Case 2
                    Clima = eColorEstado.Niebla
                
                Case 3
                    Clima = eColorEstado.FogLluvia
            
            End Select
        End If
        
    End If

    'Sea cual sea el resultado, lo mandamos
    Call SortearHorario(Clima)
    
End Sub
