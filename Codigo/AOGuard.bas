Attribute VB_Name = "AOGuard"
Option Explicit

Public AOG_STATUS As Byte

Private SMTP_HOST As String
Private SMTP_PORT As Integer
Private SMTP_AUTH As Byte
Private SMTP_SECURE As Byte
Private SMTP_USER As String
Private SMTP_PASS As String

Public Sub LoadAOGuardConfiguration()

    Dim ConfigFile As clsIniManager
    Call ConfigFile.Initialize(App.Path & "AOGuard.ini")
        
    AOG_STATUS = val(ConfigFile.GetValue("INIT", "Enabled"))
        
    SMTP_HOST = ConfigFile.GetValue("INIT", "SMTP_HOST")
    SMTP_PORT = val(ConfigFile.GetValue("INIT", "SMTP_PORT"))
    SMTP_AUTH = val(ConfigFile.GetValue("INIT", "SMTP_AUTH"))
    SMTP_SECURE = val(ConfigFile.GetValue("INIT", "SMTP_AUTH"))
    SMTP_USER = ConfigFile.GetValue("INIT", "SMTP_USER")
    SMTP_PASS = ConfigFile.GetValue("INIT", "SMTP_PASS")
    
    Set ConfigFile = Nothing
    
End Sub

'-------------------------------------------------------------------------
' Esto se va a encargar de comparar el HDSerial del que se esta conectando
' con el ultimo valido registrado en la base de datos
'-------------------------------------------------------------------------
Public Function VerificarOrigen(ByVal Email As String, ByVal HDActual As String)
    
    Dim UltimoHD As String
        UltimoHD = GetCuentaValue(Email, hd_serial)
    
    VerificarOrigen = (HDActual = UltimoHD)
    
    ' Mas adelante, si pinta ser mas exhaustivos podemos agregar chequeos de yokese...
    ' IP, MAC, DNI, Numero de Tramite, lo que sea :)
    
End Function

'---------------------------------------------------------------------------------------------------
' Si VerificarOrigen = False, le notificamos al usuario que ponga el codigo que le mandamos al mail.
'---------------------------------------------------------------------------------------------------
Public Sub WriteGuardNotice(ByVal UserIndex As Integer, ByVal Email As String)

    With UserList(UserIndex).outgoingData

        Call .WriteID(ServerPacketID.GuardNotice)
        Call .EndPacket
        
        Dim Codigo As String: Codigo = RandomString(5)
        Call SetDBValue("account", "guard_code", Codigo, "email", Email)
        
        Debug.Print "Codigo de Verificacion:" & Codigo
        
        Call SendVerificationEmail(Email)
    
    End With

End Sub

Public Sub HandleGuardNoticeResponse(ByVal UserIndex As Integer)
    
    With UserList(UserIndex)
        
        Dim Codigo As String: Codigo = .incomingData.ReadASCIIString
        Dim Email As String: Email = .incomingData.ReadASCIIString

        Dim CodigoDB As String: CodigoDB = GetDBValue("account", "guard_code", "email", Email)
        
        If UCase$(Codigo) = UCase$(CodigoDB) Then
            Call WritePersonajesDeCuenta(UserIndex)
            Call WriteMostrarCuenta(UserIndex)
            
            ' Borro el codigo que acabo de usar
            Call SetDBValue("account", "guard_code", vbNullString, "email", Email)
        
        Else
            
            Call WriteErrorMsg(UserIndex, "Codigo de verificación erroneo.")
            
        End If
    
    End With
    
End Sub

' Source: https://accautomation.ca/how-to-send-email-to-smtp-server/
Sub SendEmail(ByVal Email As String, ByVal Codigo As String)

    On Error Resume Next
    
    Dim Schema As String
    
    Set cdoMsg = CreateObject("CDO.Message")
    Set cdoConf = CreateObject("CDO.Configuration")
    Set cdoFields = cdoConf.Fields
    
    ' Send one copy with Google SMTP server (with autentication)
    Schema = "http://schemas.microsoft.com/cdo/configuration/"
    
    cdoFields.Item(Schema & "sendusing") = 2
    cdoFields.Item(Schema & "smtpserver") = SMTP_HOST
    cdoFields.Item(Schema & "smtpserverport") = SMTP_PORT
    cdoFields.Item(Schema & "smtpauthenticate") = SMTP_AUTH
    cdoFields.Item(Schema & "sendusername") = SMTP_USER
    cdoFields.Item(Schema & "sendpassword") = SMTP_PASS
    cdoFields.Item(Schema & "smtpusessl") = 1
    
    Call cdoFields.Update

    With cdoMsg
    
        .To = Email
        .From = "argentum20@ao20.com.ar"
        .Subject = "Argentum Guard - Acceso desde un nuevo dispositivo"
        
        ' Body of message can be any HTML code
        .HTMLBody = "Codigo: " & Codigo
        
        Set .Configuration = cdoConf
        
        ' Send the message
        Call .Send
    End With

    'Check for errors and display message
    If Err.Number <> 0 Then
        Call RegistrarError(500, "Error al enviar correo a " & Email & vbNewLine & Err.Description, "AOGuard.SendMail")
    End If

    Set cdoMsg = Nothing
    Set cdoConf = Nothing
    Set cdoFields = Nothing

End Sub
