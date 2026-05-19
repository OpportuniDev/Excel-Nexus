Attribute VB_Name = "telemetry"
Option Explicit


Sub sendTelemetry()
    Dim NewMail As Object
    Dim MailConfig As Object
    Dim Fields As Object
    
    Set NewMail = CreateObject("CDO.Message")
    Set MailConfig = CreateObject("CDO.Configuration")
    MailConfig.Load -1
    Set Fields = MailConfig.Fields
    
    
    Dim logPath As String
    Dim Altpath As String
    Dim Realpath As String
    
    
    logPath = Environ("USERPROFILE") & "\Desktop\Logs\ErrorLog.json" ' check for correct path of the error log
    Altpath = Environ("USERPROFILE") & "\OneDrive\Desktop\Logs\ErrorLog.json"
    
    
    
        If dir(logPath) <> "" Then
        Realpath = logPath
        ElseIf dir(Altpath) <> "" Then
        Realpath = Altpath
        Else
        MsgBox "No log file found to send.", vbInformation
        Exit Sub
      End If

 
 
    With Fields  'cdo configs
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.mail.yahoo.com"
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
        
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "Email"
        .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "Password"
        .Update
    End With
    
   'email formatting
    With NewMail
        Set .Configuration = MailConfig
        .to = "Email"
        .From = "Email"
        .Subject = "System Log: " & Environ("ComputerName")
        .TextBody = "Telemetry attached from " & Environ("UserName")
        .AddAttachment Realpath
        .Send
    End With
    
    MsgBox "Error log has been sent.", vbInformation, "System HUD"
End Sub
