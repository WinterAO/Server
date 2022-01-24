Attribute VB_Name = "modMail"
Option Explicit

Private txt_Servidor As String
Private txt_Para As String
Private txt_Puerto As String
Private txt_De As String
Private txt_Mensaje As String
Private txt_Asunto As String
Private txt_Usuario As String
Private txt_Password As String

Private Function Enviar_Mail_CDO(SerVidor_SMTP As String, _
                                 Para As String, _
                                 De As String, _
                                 Asunto As String, _
                                 Mensaje As String, _
                                 Optional Path_Adjunto As String, _
                                 Optional Puerto As String = 25, _
                                 Optional Usuario As String, _
                                 Optional Password As String, _
                                 Optional Usar_Autentificacion As Boolean = True, _
                                 Optional Usar_SSL As Boolean = True) As Boolean

        'Cambia el puntero a ocupado
        frmMain.MousePointer = vbHourglass

        ' Variable de objeto Cdo.Message
        Dim Obj_Email As CDO.Message

        ' Crea un Nuevo objeto CDO.Message
        Set Obj_Email = New CDO.Message

        ' Indica el servidor Smtp para poder enviar el Mail ( puede ser el nombre _
          del servidor o su dirección IP )
        Obj_Email.Configuration.Fields(cdoSMTPServer) = SerVidor_SMTP

        Obj_Email.Configuration.Fields(cdoSendUsingMethod) = 2

        ' Puerto. Por defecto se usa el puerto 25, en el caso de Gmail se usan los puertos _
465       o el puerto 587 ( este último me dio error )

        Obj_Email.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = CLng(Puerto)

        ' Indica el tipo de autentificación con el servidor de correo _
          El valor 0 no requiere autentificarse, el valor 1 es con autentificación
        Obj_Email.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/" & "configuration/smtpauthenticate") = Abs(Usar_Autentificacion)

        ' Tiempo máximo de espera en segundos para la conexión
        Obj_Email.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 5

        ' Configura las opciones para el login en el SMTP
        If Usar_Autentificacion Then

            ' Id de usuario del servidor Smtp ( en el caso de gmail, debe ser la dirección de correro _
              mas el @gmail.com )
            Obj_Email.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = Usuario

            ' Password de la cuenta
            Obj_Email.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = Password

            ' Indica si se usa SSL para el envío. En el caso de Gmail requiere que esté en True
            Obj_Email.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = Usar_SSL

        End If

        ' ************************************************** *******************************
        ' Estructura del mail
        '************************************************* *********************************

        ' Dirección del Destinatario
        Obj_Email.To = Para

        ' Dirección del remitente
        Obj_Email.from = De

        ' Asunto del mensaje
        Obj_Email.Subject = Asunto

        ' Cuerpo del mensaje
        Obj_Email.TextBody = Mensaje

        'Ruta del archivo adjunto

        If Path_Adjunto <> vbNullString Then
            Obj_Email.AddAttachment (Path_Adjunto)

        End If

        ' Actualiza los datos antes de enviar
        Obj_Email.Configuration.Fields.Update

        On Error Resume Next

        ' Envía el email
        Obj_Email.send

        If Err.Number = 0 Then
            Enviar_Mail_CDO = True
        Else
            MsgBox Err.description, vbCritical, " Error al enviar el amil "

        End If

        ' Descarga la referencia y deja la variable vacia (destructor)
        If Not Obj_Email Is Nothing Then
            Set Obj_Email = Nothing

        End If

        'Regresar el puntero a la normalidad
        On Error GoTo 0

        frmMain.MousePointer = vbNormal

End Function

Private Sub Enviar()

    Dim ret As Boolean

    ' Asegurarse de pasar bien los últimos dos parámetros _
      ( Si usa login y si el server usa SSL)

    ret = Enviar_Mail_CDO(txt_Servidor, txt_Para, txt_De, txt_Asunto, txt_Mensaje, , txt_Puerto, txt_Usuario, txt_Password, True, True)

    ' Si devuelve true es por que no hubo errores en el envio
    If ret Then _
        Call LogCreaciondeCuentas(txt_Mensaje)

End Sub

Public Function enviarMailVerificacion(ByVal UserName As String, ByVal Email As String, ByVal CodigoVerificacion As String)
    '***************************************************
    'Author: Lorwik
    'Last Modification: ????
    'Descripcion: Envia el codigo de activacion de la cuenta por email
    '***************************************************
    On Error GoTo ErrorHandler
    
    txt_Servidor = "smtp.gmail.com"
    txt_Para = "lorwik@gmail.com"
    txt_De = "kiwrol1992@gmail.com"
    txt_Asunto = "Activa tu cuenta en WinterAO - Codigo de verificación"
    txt_Mensaje = "Hola " & UserName & "! Acá te enviamos tu codigo de verificacion para que puedas activar tu nueva cuenta en WinterAO. CODIGO: " & CodigoVerificacion
    txt_Puerto = 465
    txt_Password = "Perolo30?"
    txt_Usuario = "kiwrol1992@gmail.com"

    Call Enviar
    
    enviarMailVerificacion = True
    
    Exit Function
    
ErrorHandler:
    Call LogDatabaseError("Error in enviarMailVerificacion: " & UserName & ". " & Err.Number & " - " & Err.description)
    enviarMailVerificacion = False
    
End Function

'Public Sub EnviarMail()
'
'    txt_Servidor = "smtp.gmail.com"
'    txt_Para = "lorwik@gmail.com"
'    txt_De = "kiwrol1992@gmail.com"
'    txt_Asunto = "Listado de fallas."
'    txt_Mensaje = "Se adjunta una lista de las fallas que presentaron las maquinas."
'    txt_Puerto = 465
'    txt_Password = "Perolo30?"
'    txt_Usuario = "kiwrol1992@gmail.com"
'
'    Call Enviar
'
'End Sub
