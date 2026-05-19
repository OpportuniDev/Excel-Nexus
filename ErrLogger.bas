Attribute VB_Name = "Errlogger"
Option Explicit

Public Sub ErrLog(errsource As String, subname As String, errnum As Long, errdesc As String, errline As Long, errtime As Date)
    Dim ErrorPath As String, filePath As String, fNum As Integer
    Dim FSO As Object
    Dim jsonstring As String
    
    
   
    jsonstring = "{""Time"": """ & errtime & """, ""Source"": """ & errsource & """,""Subroutine"": """ & subname & """, ""Num"": " & errnum & ", ""Desc"": """ & errdesc & """, ""Line"": " & errline & "}"
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    ErrorPath = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\Logs"
    
    If Not FSO.FolderExists(ErrorPath) Then
        FSO.CreateFolder ErrorPath
    End If
    
    filePath = ErrorPath & "\ErrorLog.json"
    fNum = FreeFile
    
    
    Open filePath For Append As #fNum
        Print #fNum, jsonstring & vbNewLine
    Close #fNum

    End Sub

 

