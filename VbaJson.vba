Public Function quoteStr(ByVal str As String) As String
    quoteStr = Chr(34) & str & Chr(34)
End Function

Public Function write2File(ByVal FileName As String, ByVal Content As String) ', ByVal DeviceType As String', ByVal IsSchema As Boolean)
     Dim objStream As Object
     Set objStream = CreateObject("ADODB.Stream")
     'Dim FileName As String
    
'     If IsSchema Then
'        FileName = Application.GetSaveAsFilename(DeviceType & ".Schema.json", "(*.json),*.json")
'     Else
'        FileName = Application.GetSaveAsFilename(DeviceType & ".Instance.json", "(*.json),*.json")
'     End If
     
     If FileName <> "False" Then
        With objStream
            .Type = 2
            .Charset = "UTF-8"
            .Open
            .WriteText Content
            .SaveToFile FileName, 2
        End With
    Else
        'MsgBox "Saving Failed", vbOKOnly, "Saving as Json"
    End If

    Set objStream = Nothing
End Function
Function ChangeLine(ByVal num As Integer) As String
    Dim str As String
    str = vbCr & vbLf
    If num > 0 Then
        For i = 1 To num
            str = str & vbTab
        Next
    End If
    ChangeLine = str
End Function
Function CL(num As Integer)
    If num > 0 Then
        TabNumber = TabNumber + 1
    ElseIf num = 0 Then
        TabNumber = TabNumber + 0
    ElseIf num < 0 Then
        TabNumber = TabNumber - 1
    Else
    
    End If
    CL = ChangeLine(TabNumber)
End Function
Sub AddHead(name As String)
    recordJsonStr = recordJsonStr & quoteStr(name) & " : {" & CL(1)
End Sub
Sub AddAtt(attName As String, va As String)
    recordJsonStr = recordJsonStr & quoteStr(attName) & " : " & quoteStr(va) & CL(-1)
End Sub
Sub AddAttC(attName As String, va As String)
    recordJsonStr = recordJsonStr & quoteStr(attName) & " : " & quoteStr(va) & "," & CL(0)
End Sub
Sub AddV(attName As String, va As String, AddQuo As Boolean)
    
    recordJsonStr = recordJsonStr & quoteStr(attName)
    
    If AddQuo = True Then
        recordJsonStr = recordJsonStr & " : " & quoteStr(va) & CL(-1)
    Else
        recordJsonStr = recordJsonStr & " : " & va & CL(-1)
    End If
End Sub
Sub AddVC(attName As String, va As String, AddQuo As Boolean)
    
    recordJsonStr = recordJsonStr & quoteStr(attName)
    
    If AddQuo = True Then
        recordJsonStr = recordJsonStr & " : " & quoteStr(va) & "," & CL(0)
    Else
        recordJsonStr = recordJsonStr & " : " & va & "," & CL(0)
    End If
End Sub
Sub AddInstanceV(attName As String, va As String, ty As String)
    If InStr(0, ty, "ULONG", vbTextCompare) > 0 Or InStr(0, ty, "UWORD", vbTextCompare) > 0 Or ty = "integer" Or ty = "Bit" Then
        Call AddV(attName, va, False)
    Else
        Call AddV(attName, va, True)
    End If
End Sub
Sub AddInstanceVC(attName As String, va As String, ty As String)
    If InStr(0, ty, "ULONG", vbTextCompare) > 0 Or InStr(0, ty, "UWORD", vbTextCompare) > 0 Or ty = "integer" Or ty = "Bit" Then
        Call AddVC(attName, va, False)
    Else
        Call AddVC(attName, va, True)
    End If
End Sub
Sub ShutDown()
    recordJsonStr = recordJsonStr & "}" & CL(-1)
End Sub
Sub ShutDownC()
    recordJsonStr = recordJsonStr & "}," & CL(0)
End Sub
Function AddArray(name As String, thisArray())
    Dim i As Integer
    
    recordJsonStr = recordJsonStr & quoteStr(name) & " : [" & CL(1)
        
    For i = 0 To UBound(thisArray)
        If i = UBound(thisArray) Then
            recordJsonStr = recordJsonStr & quoteStr(thisArray(i)) & CL(-1)
        Else
            recordJsonStr = recordJsonStr & quoteStr(thisArray(i)) & "," & CL(0)
        End If
    Next
    Call ArrayShutDown
End Function
Function AddArrayC(name As String, thisArray())
    Dim i As Integer
    
    recordJsonStr = recordJsonStr & quoteStr(name) & " : [" & CL(1)
        
    For i = 0 To UBound(thisArray)
        If i = UBound(thisArray) Then
            recordJsonStr = recordJsonStr & quoteStr(thisArray(i)) & CL(-1)
        Else
            recordJsonStr = recordJsonStr & quoteStr(thisArray(i)) & "," & CL(0)
        End If
    Next
    Call ArrayShutDownC
End Function
Sub ArrayShutDown()
    recordJsonStr = recordJsonStr & "]" & CL(-1)
End Sub
Sub ArrayShutDownC()
    recordJsonStr = recordJsonStr & "]," & CL(0)
End Sub

Sub ArrayShutDown()
    recordJsonStr = recordJsonStr & "]" & CL(-1)
End Sub
Sub ArrayShutDownC()
    recordJsonStr = recordJsonStr & "]," & CL(0)
End Sub