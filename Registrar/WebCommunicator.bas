Attribute VB_Name = "WebCommunicator"
Option Explicit
' we set this to true whil a connection is established
Public blnConnected As Boolean

Public Sub sendRequest(ByRef sckTarget As Winsock, endpoint As String, formData As Dictionary, strMethod As String)
    Dim eUrl As URL
    Dim strData As String
    Dim strPostData As String
    Dim strHeaders As String
    
    Dim strHTTP As String
    Dim x As Integer
    
    strPostData = ""
    strHeaders = ""
    If blnConnected Then Exit Sub
    
    ' get the url
    eUrl = ExtractUrl(ipaddress & endpoint)
    
    If eUrl.Host = vbNullString Then
        MsgBox "Invalid Host", vbCritical, "ERROR"
    
        Exit Sub
    End If
    
     ' configure winsock
    sckTarget.RemoteHost = eUrl.Host
    
    If eUrl.Scheme = "http" Then
        If eUrl.Port > 0 Then
            sckTarget.RemotePort = eUrl.Port
        Else
            sckTarget.RemotePort = 80
        End If
    ElseIf eUrl.Scheme = vbNullString Then
        sckTarget.RemotePort = 80
    Else
        MsgBox "Invalid protocol schema"
    End If
    
    ' build encoded data the data is url encoded in the form
    ' var1=value&var2=value
    strData = ""
    Dim key
    For Each key In formData.keys
        strData = strData & key & "=" & formData(key) & "&"
    Next
                            
    If eUrl.Query <> vbNullString Then
        eUrl.URI = eUrl.URI & "?" & eUrl.Query
    End If
    
    ' check if any variables were supplied
    If strData <> vbNullString Then
        strData = Left(strData, Len(strData) - 1)
    
        If strMethod = "GET" Then
            ' if this is a GET request then the URL encoded data
            ' is appended to the URI with a ?
            If eUrl.Query <> vbNullString Then
                eUrl.URI = eUrl.URI & "&" & strData
            Else
                eUrl.URI = eUrl.URI & "?" & strData
            End If
        Else
            ' if it is a post request, the data is appended to the
            ' body of the HTTP request and the headers Content-Type
            ' and Content-Length added

            strPostData = strData
            strHeaders = "Content-Type: application/x-www-form-urlencoded" & vbCrLf & _
                         "Content-Length: " & Len(strPostData) & vbCrLf
                         
        End If
    End If
    
    ' clear the old HTTP response
    Dim response As String
    
    ' build the HTTP request in the form
    '
    ' {REQ METHOD} URI HTTP/1.0
    ' Host: {host}
    ' {headers}
    '
    ' {post data}
    strHTTP = strMethod & " " & eUrl.URI & " HTTP/1.0" & vbCrLf
    strHTTP = strHTTP & "Host: " & eUrl.Host & vbCrLf
    strHTTP = strHTTP & strHeaders
    strHTTP = strHTTP & vbCrLf
    strHTTP = strHTTP & strPostData

    response = strHTTP
    sckTarget.Connect
    
    ' wait for a connection
    While Not blnConnected
        DoEvents
    Wend
    
    ' send the HTTP request
    sckTarget.SendData strHTTP
End Sub
