Attribute VB_Name = "url_helper"
Option Explicit

Type URL
    Scheme As String
    Host As String
    Port As Long
    URI As String
    Query As String
End Type
    

' returns as type URL from a string
Function ExtractUrl(ByVal strUrl As String) As URL
    Dim intPos1 As Integer
    Dim intPos2 As Integer
    
    Dim retURL As URL
    
    '1 look for a scheme it ends with ://
    intPos1 = InStr(strUrl, "://")
    
    If intPos1 > 0 Then
        retURL.Scheme = Mid(strUrl, 1, intPos1 - 1)
        strUrl = Mid(strUrl, intPos1 + 3)
    End If
        
    '2 look for a port
    intPos1 = InStr(strUrl, ":")
    intPos2 = InStr(strUrl, "/")
    
    If intPos1 > 0 And intPos1 < intPos2 Then
        ' a port is specified
        retURL.Host = Mid(strUrl, 1, intPos1 - 1)
        
        If (IsNumeric(Mid(strUrl, intPos1 + 1, intPos2 - intPos1 - 1))) Then
                retURL.Port = CInt(Mid(strUrl, intPos1 + 1, intPos2 - intPos1 - 1))
        End If
    ElseIf intPos2 > 0 Then
        retURL.Host = Mid(strUrl, 1, intPos2 - 1)
    Else
        retURL.Host = strUrl
        retURL.URI = "/"
        
        ExtractUrl = retURL
        Exit Function
    End If
    
    strUrl = Mid(strUrl, intPos2)
    
    ' find a question mark ?
    intPos1 = InStr(strUrl, "?")
    
    If intPos1 > 0 Then
        retURL.URI = Mid(strUrl, 1, intPos1 - 1)
        retURL.Query = Mid(strUrl, intPos1 + 1)
    Else
        retURL.URI = strUrl
    End If
    
    ExtractUrl = retURL
End Function

' url encodes a string
Function URLEncode(ByVal str As String) As String
        Dim intLen As Integer
        Dim x As Integer
        Dim curChar As Long
                Dim newStr As String
                intLen = Len(str)
        newStr = ""
                        For x = 1 To intLen
            curChar = Asc(Mid$(str, x, 1))
            
            If (curChar < 48 Or curChar > 57) And _
                (curChar < 65 Or curChar > 90) And _
                (curChar < 97 Or curChar > 122) Then
                                newStr = newStr & "%" & Hex(curChar)
            Else
                newStr = newStr & Chr(curChar)
            End If
        Next x
        
        URLEncode = newStr
End Function

' decodes a url encoded string
Function UrlDecode(ByVal str As String) As String
        Dim intLen As Integer
        Dim x As Integer
        Dim curChar As String * 1
        Dim strCode As String * 2
        
        Dim newStr As String
        
        intLen = Len(str)
        newStr = ""
        
        For x = 1 To intLen
            curChar = Mid$(str, x, 1)
            
            If curChar = "%" Then
                strCode = "&h" & Mid$(str, x + 1, 2)
                
                If IsNumeric(strCode) Then
                    curChar = Chr(Int(strCode))
                Else
                    curChar = ""
                End If
                                x = x + 2
            End If
            
            newStr = newStr & curChar
        Next x
        
        UrlDecode = newStr
End Function

Function getJSONFromResponse(response As String) As String
    Dim JSON As String
    JSON = Mid$(response, InStr(1, response, "{"))
    getJSONFromResponse = JSON
End Function
