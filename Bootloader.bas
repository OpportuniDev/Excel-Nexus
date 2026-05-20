Attribute VB_Name = "Bootloader"
Option Explicit

Public Sub BootLoad()

 

10    Dim http As Object
20    Dim fileObj As Object
30    Dim fileContent As String
40    Dim tempPath As String
50 Dim tempFile As String
60 Dim Updates As Collection
70   Set Updates = New Collection
80     tempPath = Environ("TEMP") & "\"
90  Dim testProject As Object
        
100        On Error Resume Next
120    Set testProject = ThisWorkbook.VBProject
    
130   If Err.Number <> 0 Then
140        MsgBox "SYSTEM ACCESS REQUIRED:" & vbCrLf & vbCrLf & _
               "This updater needs permission to modify VBA modules." & vbCrLf & _
               "Please perform these 3 steps:" & vbCrLf & _
               "1. Go to File > Options > Trust Center > Trust Center Settings." & vbCrLf & _
               "2. Select 'Macro Settings'." & vbCrLf & _
               "3. Check 'Trust access to the VBA project object model'.", _
               vbCritical, "Security Blocked"
150       Exit Sub
160    End If
170    On Error GoTo 0
    


   


180  Dim fileURL As String
190    'fileURL = "host url here"
    
200    Set http = CreateObject("MSXML2.XMLHTTP")
210    http.Open "GET", fileURL, False
220    http.Send
    
230    If http.Status = 200 Then
        
240        fileContent = http.responseText
    
250   Set fileObj = New ClsFileModule
        
        
260        Dim urlParts() As String
270         Dim headerName As String
280         Dim nameParts As Variant
290            Dim nameFinal As Variant
300            Dim nameAlmost As String
310            On Error GoTo errorhandler
            

320        urlParts = Split(fileURL, "/")
330        headerName = http.getResponseHeader("Content-Disposition")
340            nameParts = Split(headerName, "=")
350           nameAlmost = nameParts(UBound(nameParts))
360           nameFinal = Split(nameAlmost, ".")
370           fileObj.Name = nameFinal(LBound(nameFinal))
380 On Error GoTo errorhandler

        
        
390        fileObj.Code = fileContent
400         Updates.Add fileObj
410    Else
420        MsgBox "Failed to download file: " & http.Status & " - " & http.statusText
430        Set fileObj = Nothing
440    End If

450 On Error GoTo errorhandler

460 Dim modName As Object
470     For Each fileObj In Updates
480             For Each modName In ThisWorkbook.VBProject.VBComponents
490                 If LCase(fileObj.Name) = LCase(modName.Name) Then
500           ThisWorkbook.VBProject.VBComponents.Remove modName
510             Exit For
520         End If
530     Next modName
540 Next fileObj
550 On Error GoTo errorhandler

560 For Each fileObj In Updates

570    tempFile = tempPath & fileObj.Name
580    Dim fNum As Integer
590     fNum = FreeFile
600        Open tempFile For Output As #fNum
610            Print #fNum, fileObj.Code
620        Close #fNum
630    ThisWorkbook.VBProject.VBComponents.Import tempFile
        
640        Kill tempFile

650 Next fileObj
655 Application.OnKey "^u", "ClientUpdater"
660 On Error GoTo errorhandler
670 Exit Sub
   
680 errorhandler:
690   Dim retrycount As Long
700    Dim errline As Long
710   Dim errtime As Date
715   Dim subname As String
720    errline = Erl
730            errtime = Now
735                 subname = "Bootloader"
740                    retrycount = retrycount + 1
750                    If retrycount < 3 Then Resume
760
770                If retrycount = 3 Then
780            Call ErrLog(Err.Source, subname, Err.Number, Err.Description, errline, errtime)
790         MsgBox "An error has occured, please notify administration."
800         Application.ScreenUpdating = True
810    Application.EnableEvents = True
820 End If
830 Exit Sub

End Sub



