<SCRIPT LANGUAGE="VBSCRIPT" RUNAT="SERVER">

' Status codes
' 0 - Init
' 1 - Uploading
' 2 - Writing
' 3 - Done

Class UploadProgress
  
  Private m_UploadId
  
  Public Property Let UploadId(value)
    m_UploadId = CStr(value)
    If Not UploadIdExists() Then AddUploadId
  End Property
  
  Public Property Get UploadId()
    UploadId = m_UploadId
  End Property
  
  Public Property Let TimeOut(value)
    SetValue "TimeOut", CLng(value)
  End Property
  
  Public Property Get TimeOut()
    TimeOut = CLng(GetValueEx("TimeOut", 600))
  End Property
  
  Public Property Let TotalBytes(value)
    SetValue "TotalBytes", CLng(value)
  End Property
  
  Public Property Get TotalBytes()
    TotalBytes = CLng(GetValueEx("TotalBytes", 0))
  End Property
  
  Public Property Let UploadedBytes(value)
    SetValue "UploadedBytes", CLng(value)
  End Property
  
  Public Property Get UploadedBytes()
    UploadedBytes = CLng(GetValueEx("UploadedBytes", 0))
  End Property
  
  Public Property Let Status(value)
    SetValue "Status", CInt(value)
  End Property
  
  Public Property Get Status()
    Status = CInt(GetValueEx("Status", 0))
  End Property
  
  Public Property Get LastFile()
    Dim noFiles
    noFiles = GetNumberOf("File")
    If noFiles > 0 Then
      LastFile = GetValue("File" & noFiles)
    Else
      LastFile = ""
    End If
  End Property
  
  Public Property Get LastError()
    Dim noErrors
    noErrors = GetNumberOf("Error")
    If noErrors > 0 Then
      LastError = GetValue("Error" & noErrors)
    Else
      LastError = ""
    End If
  End Property
  
  Private Sub Class_Initialize()
    CleanUp
    m_UploadId = ""
  End Sub

  Private Sub Class_Terminate()
  End Sub
  
  Public Sub SetValue(name, value)
    If m_UploadId <> "" Then
      Application.Lock
      Application("PU3_" & m_UploadId & "." & name) = value
      Application("PU3_" & m_UploadId & ".LastUpdate") = Now()
      Application.UnLock
    End If
  End Sub
  
  Public Function GetValue(name)
    GetValue = Application("PU3_" & m_UploadId & "." & name)
  End Function
  
  Public Function GetValueEx(name, def)
    GetValueEx = GetValue(name)
    If GetValueEx = "" Then GetValueEx = def
  End Function
  
  Public Sub Remove(name)
    Application.Contents.Remove("PU3_" & m_UploadId & "." & name)
  End Sub
  
  Public Sub RemoveAll()
    Dim toDelete, content, toDeleteArr, i
    toDelete = ""
    For Each content In Application.Contents
      If InStr(content, ".") > 0 Then
        If Left(content, InStr(content, ".") - 1) = "PU3_" & m_UploadId Then
          toDelete = toDelete & "|" & content
        End If
      End If
    Next
    toDeleteArr = Split(toDelete, "|")
    For i = 1 To UBound(toDeleteArr)
      Application.Contents.Remove(toDeleteArr(i))
    Next
  End Sub
  
  Public Sub AddFile(fileName)
    Dim nr
    If m_UploadId <> "" Then
      nr = GetNumberOf("File") + 1
      SetValue "File" & nr, fileName
      SetValue "File" & nr & "Status", "Init"
      SetValue "File" & nr & "Error", ""
    End If
  End Sub
  
  Public Sub SetFileInfo(name, value)
    Dim nr
    If m_UploadId <> "" Then
      nr = GetNumberOf("File")
      SetValue "File" & nr & name, value
    End If
  End Sub
  
  Public Sub AddGlobalError(errorString)
    Dim nr
    If m_UploadId <> "" Then
      nr = GetNumberOf("Error") + 1
      SetValue "Error" & nr, errorString
    End If
  End Sub
  
  Public Function GetNumberOf(name)
    Dim nr
    nr = 0
    While GetValue(name & (nr+1)) <> ""
      nr = nr+1
    Wend
    GetNumberOf = nr
  End Function
  
  Private Function UploadIdExists()
    If Application("PU3Uploads") <> "" Then
      Dim uploads, i
      uploads = Split(Application("PU3Uploads"), "|")
      For i = 1 To UBound(uploads)
        If uploads(i) = m_UploadId Then
          UploadIdExists = True
          Exit Function
        End If
      Next
    End If
    UploadIdExists = False
  End Function
  
  Private Sub AddUploadId()
    Application("PU3Uploads") = Application("PU3Uploads") & "|" & m_UploadId
    SetValue "DateTime", Now()
  End Sub
  
  Private Sub CleanUp()
    Dim curUploadId, uploads, i, theLastUpdate, theTimeOut
    If Application("PU3Uploads") <> "" Then
      uploads = Split(Application("PU3Uploads"), "|")
      For i = 1 To UBound(uploads)
        m_UploadId = uploads(i)
        theLastUpdate = CDate(GetValue("LastUpdate"))
        theTimeOut = CLng(GetValueEx("TimeOut", 600))
        If DateDiff("s", theLastUpdate, Now()) > theTimeOut Then
          RemoveAll
          Application("PU3Uploads") = Replace(Application("PU3Uploads"), "|" & m_UploadId, "")
        End If
      Next
    End If
  End Sub
  
End Class

Class Translation
  
  Private m_Xml
  
  Private Sub Class_Initialize()
  End Sub
  
  Private Sub Class_Terminate()
    Set m_Xml = Nothing
  End Sub
  
  Public Sub Load(file)
    On Error Resume Next
    Set m_Xml = Server.CreateObject("Msxml2.DOMDocument")
    m_Xml.async = False
    m_Xml.resolveExternals = False
    m_Xml.load file
    m_Xml.setProperty "SelectionLanguage", "XPath"
    If Err.Number <> 0 Then
      On Error GoTo 0
      Err.Raise 2, "Translation", "Error loading localization file."
    End If
    On Error GoTo 0
  End Sub
  
  Public Function Value(resname)
    Dim sourceNode
    
    If Not IsObject(m_Xml) Then
      Value = "Translation for " & resname & " not found."
      Exit Function
    End If
    
    Set sourceNode = m_Xml.selectSingleNode("//trans-unit[@resname='" & resname & "']")
    
    If sourceNode Is Nothing Then
      Value = "Translation for " & resname & " not found."
      Exit Function
    End If
    
    Value = sourceNode.text
  End Function
  
  Public Function ValueEx(resname, params)
    Dim sourceNode, i
    
    If Not IsObject(m_Xml) Then
      ValueEx = "Translation for " & resname & " not found."
      Exit Function
    End If
    
    Set sourceNode = m_Xml.selectSingleNode("//trans-unit[@resname='" & resname & "']")
    
    If SourceNode Is Nothing Then
      ValueEx = "Translation for " & resname & " not found."
      Exit Function
    End If
    
    ValueEx = sourceNode.text
    If IsArray(params) Then
      For i = 0 To UBound(params)
        ValueEx = Replace(ValueEx, "%"&(i+1), params(i))
      Next
    End If
  End Function
  
  Public Sub Write(resname)
    Response.Write Value(resname)
  End Sub
  
  Public Sub WriteEx(resname, params)
    Response.Write ValueEx(resname, params)
  End Sub
  
End Class

</SCRIPT>