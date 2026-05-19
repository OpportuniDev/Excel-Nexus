Attribute VB_Name = "MasterUpdater"
Option Explicit

Sub MasterUpdater()
Dim FilePicker As Object
Dim http As Object
Dim vrtselectedItem As Variant
Dim boundary As String
Dim CRLF As String
Dim body As String
Dim fileContent() As Byte
Dim i As Long
Dim fileNum As Integer
Dim fileName As String
3
5 Set FilePicker = Application.FileDialog(3)
         'use to pick the file
10    With FilePicker
20        .AllowMultiSelect = True
30     If .Show = -1 Then
40        For Each vrtselectedItem In .SelectedItems
50
60
70            fileName = dir(vrtselectedItem)
            
            
           ' load the file into byte array
80            fileNum = FreeFile
90            Open vrtselectedItem For Binary As #fileNum
100                ReDim fileContent(LOF(fileNum) - 1)
110                Get #fileNum, , fileContent
120            Close #fileNum

            
130            boundary = "----WebKitFormBoundary7MA4YWxkTrZu0gW"
140            CRLF = vbCrLf

            ' build the post body
150            body = "--" & boundary & CRLF
160            body = body & "Content-Disposition: form-data; name=""file""; filename=""" & fileName & """" & CRLF
170            body = body & "Content-Type: application/octet-stream" & CRLF & CRLF

            ' send the file content as bytes
            
180            Dim headerBytes() As Byte
190            headerBytes = StrConv(body, vbFromUnicode)

            
200            Dim closing() As Byte
210            closing = StrConv(CRLF & "--" & boundary & "--" & CRLF, vbFromUnicode)

220             Dim sendBytes() As Byte
225             Dim vsendBytes As Variant
230            ReDim sendBytes(LBound(headerBytes) To UBound(headerBytes) + UBound(fileContent) + 1 + UBound(closing) + 1)

            
240            For i = LBound(headerBytes) To UBound(headerBytes)
250                sendBytes(i) = headerBytes(i)
260            Next i

            
270            For i = LBound(fileContent) To UBound(fileContent)
280                sendBytes(UBound(headerBytes) + i + 1) = fileContent(i)
290            Next i

           
300            For i = LBound(closing) To UBound(closing)
310                sendBytes(UBound(headerBytes) + UBound(fileContent) + 2 + i) = closing(i)
320            Next i

            
330            Set http = CreateObject("MSXML2.XMLHTTP")
340            http.Open "POST", "webaddress", False ' - add the hosting server here
350            http.setRequestHeader "Content-Type", "multipart/form-data; boundary=" & boundary
360              vsendBytes = sendBytes
362                http.Send vsendBytes


            
370
380        Next vrtselectedItem
390    End If
400 End With
End Sub
