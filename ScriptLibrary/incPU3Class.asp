<SCRIPT LANGUAGE="VBSCRIPT" RUNAT="SERVER">
'------------------------------------------
' Pure ASP Upload 3
' Copyright 2006 (c) DMXzone
' Version: 3.0.7
'------------------------------------------

' Constants for ConflictHandling
Const puConflictOverwrite = "over"
Const puConflictIgnore = "skip"
Const puConflictUnique = "uniq"
Const puConflictError = "error"

' Constants for StoreType
Const puStoreFile = "file"
Const puStorePath = "path"

Class PureUpload

  Private m_Debug
  Private m_UploadRequest()
  Private m_Count
  Private m_UploadDone
  Private m_Progress
  Private m_UploadId
  Private m_Path
  Private m_TimeOut
  Private m_MaxSize
  Private m_MaxFileSize
  Private m_AllowedExtensions
  Private m_ConflictHandling
  Private m_LastWrite
  Private m_KeepInMemory
  Private m_MinWidth
  Private m_MinHeight
  Private m_MaxWidth
  Private m_MaxHeight
  Private m_StoreType
  Private m_Required
  Private m_TempFolder
  Private m_CharSetMap
  Private m_CharSet
  Private m_ProgressTemplate
  Private m_ProgressWidth
  Private m_ProgressHeight
  Private m_ScriptLibrary
  Private m_LangFile
  Private m_Translate
  Private m_RedirectUrl
  Private m_HaltOnErrors
  Private m_RaiseErrors
  
  Private Util

  Public Property Let Debug(boolDebug)
    m_Debug = CBool(boolDebug)
  End Property
  
  Public Property Let ScriptLibrary(strPath)
    m_ScriptLibrary = Trim(CStr(strPath))
	' Set language
	m_LangFile = GetLangFile()
	m_Translate.Load(Util.GetPhysicalPath(m_LangFile))
  End Property
  
  Public Property Get ScriptLibrary()
    ScriptLibrary = m_ScriptLibrary
  End Property
  
  Public Property Let UploadFolder(strPath)
    m_Path = Trim(CStr(strPath))
    pau_thePath = m_Path
  End Property  
  
  Public Property Get UploadFolder()
    If InStr(m_Path, """") > 0 Then
      UploadFolder = Eval(m_Path)
    Else
      UploadFolder = m_Path
    End If
  End Property
  
  Public Property Let TimeOut(lngTimeOut)
    m_TimeOut = CLng(lngTimeOut)
  End Property  

  Public Property Get TimeOut()
    TimeOut = m_TimeOut
  End Property
  
  Public Property Let CharSet(strCharSet)
    m_CharSet = CStr(strCharSet)
  End Property
  
  Public Property Get CharSet()
   CharSet = m_CharSet
  End Property
  
  Public Property Let CodePage(lngCodePage)
    Dim i
    For i = 0 To UBound(m_CharSetMap) Step 3 
      If CLng(lngCodePage) = CLng(m_CharSetMap(i+2)) Then
        m_CharSet = m_CharSetMap(i+1)
        Exit For
      End If
    Next
  End Property
  
  Public Property Get CodePage()
    Dim i
    CodePage = 0
    For i = 0 To UBound(m_CharSetMap) Step 3 
      If m_CharSet = m_CharSetMap(i+1) Then
        CodePage = CLng(m_CharSetMap(i+2))
        Exit For
      End If
    Next
  End Property

  Public Property Get Done()
    Done = m_UploadDone
  End Property

  Public Property Let MaxSize(lngMaxSize)
    m_MaxSize = CLng(lngMaxSize*1024)
  End Property  

  Public Property Get MaxSize()
    MaxSize = m_MaxSize/1024
  End Property

  Public Property Let MaxFileSize(lngMaxSize)
    m_MaxFileSize = CLng(lngMaxSize*1024)
  End Property  

  Public Property Get MaxFileSize()
    MaxFileSize = m_MaxFileSize/1024
  End Property

  Public Property Let AllowedExtensions(strExtensions)
    m_AllowedExtensions = Trim(CStr(strExtensions))
  End Property  

  Public Property Get AllowedExtensions()
    AllowedExtensions = m_AllowedExtensions
  End Property
  
  Public Property Let KeepInMemory(boolMemory)
    m_KeepInMemory = CBool(boolMemory)
  End Property
  
  Public Property Get KeepInMemory()
    KeepInMemory = m_KeepInMemory
  End Property
  
  Public Property Let MinWidth(intWidth)
    m_MinWidth = CLng(Abs(intWidth))
  End Property
  
  Public Property Get MinWidth()
    MinWidth = m_MinWidth
  End Property
  
  Public Property Let MinHeight(intHeight)
    m_MinHeight = CLng(Abs(intHeight))
  End Property
  
  Public Property Get MinHeight()
    MinHeight = m_MinHeight
  End Property
  
  Public Property Let MaxWidth(intWidth)
    m_MaxWidth = CLng(Abs(intWidth))
  End Property
  
  Public Property Get MaxWidth()
    MaxWidth = m_MaxWidth
  End Property
  
  Public Property Let MaxHeight(intHeight)
    m_MaxHeight = CLng(Abs(intHeight))
  End Property
  
  Public Property Get MaxHeight()
    MaxHeight = m_MaxHeight
  End Property
  
  Public Property Let ConflictHandling(strOption)
    m_ConflictHandling = Trim(CStr(strOption))
    pau_nameConflict = m_ConflictHandling
    If m_ConflictHandling <> puConflictIgnore And m_ConflictHandling <> puConflictUnique And m_ConflictHandling <> puConflictError And m_ConflictHandling <> puConflictOverwrite Then
      Util.ThrowError 4, m_Translate.ValueEx("PU3_ERR_VALUE", Array("ConflictHandling")), True
    End If
  End Property

  Public Property Get ConflictHandling()
    ConflictHandling = m_ConflictHandling
  End Property
  
  Public Property Let StoreType(strOption)
    m_StoreType = Trim(CStr(strOption))
    If m_StoreType <> puStoreFile And m_StoreType <> puStorePath Then
      Util.ThrowError 4, m_Translate.ValueEx("PU3_ERR_VALUE", Array("StoreType")), True
    End If
  End Property
  
  Public Property Get StoreType()
    StoreType = m_StoreType
  End Property
  
  Public Property Let Required(boolReq)
    m_Required = CBool(boolReq)
  End Property
  
  Public Property Get Required()
    Required = m_Required
  End Property
  
  Public Property Let ProgressTemplate(strTemplate)
    m_ProgressTemplate = CStr(strTemplate)
  End Property
  
  Public Property Get ProgressTemplate()
    ProgressTemplate = m_ProgressTemplate
  End Property
  
  Public Property Let ProgressWidth(intWidth)
    m_ProgressWidth = CLng(Abs(intWidth))
  End Property
  
  Public Property Get ProgressWidth()
    ProgressWidth = m_ProgressWidth
  End Property
  
  Public Property Let ProgressHeight(intHeight)
    m_ProgressHeight = CLng(Abs(intHeight))
  End Property
  
  Public Property Get ProgressHeight()
    ProgressHeight = m_ProgressHeight
  End Property
  
  Public Property Let TempFolder(strFolder)
    m_TempFolder = Util.GetPhysicalPath(Trim(CStr(strFolder)))
  End Property
  
  Public Property Get TempFolder()
    TempFolder = m_TempFolder
  End Property
  
  Public Property Let RedirectUrl(strUrl)
    m_RedirectUrl = Trim(CStr(strUrl))
  End Property
  
  Public Property Get RedirectUrl()
    RedirectUrl = m_RedirectUrl
  End Property
  
  Public Property Let HaltOnErrors(boolHalt)
    m_HaltOnErrors = CBool(boolHalt)
  End Property
  
  Public Property Get HaltOnErrors()
    HaltOnErrors = m_HaltOnErrors
  End Property
  
  Public Property Let RaiseErrors(boolRaise)
    m_RaiseErrors = boolRaise
    Util.RaiseErrors = m_RaiseErrors
  End Property
  
  Public Property Get RaiseErrors()
    RaiseErrors = m_RaiseErrors
  End Property

  Public Property Get TotalBytes()
    TotalBytes = Request.TotalBytes
  End Property

  Public Property Get Count()
    Count = m_Count
  End Property

  Public Default Property Get Fields(FieldName)
    Dim Index

    If Not m_UploadDone Then
      Set Fields = New UploadControl
      Exit Property
    End If

    ' If a number was passed
    If IsNumeric(FieldName) Then
      Index = CLng(FieldName)

      ' If programmer requested an invalid number
      If Index > m_Count - 1 Or Index < 0 Then
        ' Give back empty
        Set Fields = New UploadControl
        Exit Property
      End If

      ' Return the field class for the index specified
      Set Fields = m_UploadRequest(Index)
      Exit Property
    Else
      ' convert name to lowercase
      FieldName = LCase(FieldName)
      
      ' Loop through each field
      For Index = 0 To m_Count - 1
        ' If name matches current fields name in lowercase
        If LCase(m_UploadRequest(Index).Name) = FieldName Then
          ' Return Field Class
          Set Fields = m_UploadRequest(Index)
          Exit Property
        End If
      Next
    End If
    
    ' Give back empty
    Set Fields = New UploadControl
  End Property

  Private Sub Class_Initialize()
    ReDim m_UploadRequest(-1)
    m_Debug = False
    m_Path = ""
    m_TimeOut = 600
    m_UploadDone = False
	m_MaxSize = 0
	m_AllowedExtensions = ""
	m_KeepInMemory = False
	m_MinWidth = 0
	m_MinHeight = 0
	m_MaxWidth = 0
	m_MaxHeight = 0
	m_ConflictHandling = puConflictOverwrite
	m_StoreType = puStoreFile
	m_Required = False
    m_Count = 0
	pau_thePath = ""
	pau_saveWidth = ""
	pau_saveHeight = ""
	pau_nameConflict = puConflictOverwrite
	m_RedirectUrl = ""
	m_HaltOnErrors = True
	m_RaiseErrors = False
	
	Set m_Translate = New Translation
	Set Util = New PureUploadUtils
	Util.Translation = m_Translate

    On Error Resume Next
    Set UploadRequest = Server.CreateObject("Scripting.Dictionary")
    If Err.Number <> 0 Then
      On Error GoTo 0
      Util.ThrowError 18, m_Translate.ValueEx("PU3_ERR_SCRIPT", Array("Scripting.Dictionary")), True
    End If
    On Error GoTo 0
	
	m_CharSet = ""
	initCharsetMap
	
	' Detect codepage
	On Error Resume Next
	CodePage = Response.CodePage
	If Err.Number <> 0 Then
	  CodePage = Session.CodePage
	End If
	On Error GoTo 0
	
	' Get/Set the UploadId
    UploadQueryString = Request.QueryString
	If Request.QueryString("UploadId") <> "" Then
	  m_UploadId = Request.QueryString("UploadId")
	Else
	  m_UploadId = Session.SessionID
      If UploadQueryString <> "" Then
        UploadQueryString = UploadQueryString & "&GP_upload=true&UploadId=" & m_UploadId
      Else
        UploadQueryString = "GP_upload=true&UploadId=" & m_UploadId
      End If
    End If
    
    DMX_uploadAction = Request.ServerVariables("URL") & "?" & UploadQueryString
    
    ' Check if form is posted
    If LCase(Request.ServerVariables("Request_Method")) <> "post" Then
      Exit Sub
    End If
  End Sub

  Private Sub Class_Terminate()
    'Delete temp files if they still exist
    Set Util = Nothing
    Set m_Translate = Nothing
    Dim fso, Index
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    For Index = 0 To m_Count - 1
      If fso.FileExists(m_UploadRequest(Index).TempFileName) Then
        fso.DeleteFile(m_UploadRequest(Index).TempFileName)
      End If
    Next
    Set fso = Nothing
  End Sub
  
  Private Function GetLangFile()
    Dim fso, lang, i
    
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    lang = Split(Request.ServerVariables("HTTP_ACCEPT_LANGUAGE"), ",")
    GetLangFile = ""
    For i = 0 To UBound(lang)
      If Len(lang(i)) > 2 Then lang(i) = Left(lang(i), 2)
      GetLangFile = m_ScriptLibrary & "/localization/PureUpload3_" & lang(i) & ".xml"
      If fso.FileExists(Util.GetPhysicalPath(GetLangFile)) Then Exit For
      GetLangFile = ""
    Next
    
    If GetLangFile = "" Then
      GetLangFile = m_ScriptLibrary & "/Localization/PureUpload3_en.xml"
    End If
    
    Set fso = Nothing
  End Function
  
  Public Function AddField(strName)
    Dim Control
    Set Control = New UploadControl
    Control.Translation = m_Translate
    Control.RaiseErrors = m_RaiseErrors
    Control.Name = strName
    ReDim Preserve m_UploadRequest(m_Count)
    Set m_UploadRequest(m_Count) = Control
    m_Count = m_Count + 1
	Set AddField = Control
  End Function
  
  Public Function FlashVars()
    FlashVars = "url=" & Server.URLEncode(DMX_uploadAction & "&MM_insert=FlashUpload") & _
                "&id=" & Server.URLEncode(m_UploadId)
    If m_ProgressTemplate <> "" Then
      FlashVars = FlashVars & "&t=" & Server.URLEncode(m_ProgressTemplate & "?UploadId=" & m_UploadId) & _
                              "&w=" & Server.URLEncode(m_ProgressWidth) & _
                              "&h=" & Server.URLEncode(m_ProgressHeight)
    End If
    If m_AllowedExtensions <> "" Then
      FlashVars = FlashVars & "&allowedExtensions=" & Server.URLEncode(m_AllowedExtensions)
    End If
    If m_MaxSize > 0 Then
      FlashVars = FlashVars & "&maxSize=" & Server.URLEncode(m_MaxSize/1024)
    End If
    If m_MaxFileSize > 0 Then
      FlashVars = FlashVars & "&maxFileSize=" & Server.URLEncode(m_MaxFileSize/1024)
    End If
    If m_LangFile <> "" Then
      FlashVars = FlashVars & "&langUrl=" & Server.URLEncode(m_LangFile)
    End If
    FlashVars = FlashVars & "&progressUrl=" & Server.URLEncode(m_ScriptLibrary & "/PU3Progress.asp?UploadId=" & m_UploadId)
    If m_RedirectUrl <> "" Then
      FlashVars = FlashVars & "&redirectUrl=" & Server.URLEncode(m_RedirectUrl)
    End If
  End Function
  
  Public Function generateScriptCode()
    Dim jsString
    jsString = "var PU3_ERR_EXTENSION = '" & Replace(m_Translate.Value("PU3_ERR_EXTENSION"), "'", "\'") & "';" & VbCrLf & _
               "var PU3_ERR_REQUIRED = '" & Replace(m_Translate.Value("PU3_ERR_REQUIRED"), "'", "\'") & "';"
    generateScriptCode = jsString
  End Function
  
  Public Function SubmitCode()
    Dim jsString, i, Control
    ' Generate JavaScript for onSubmit
    jsString = "validateForm(this, '" & m_AllowedExtensions & "', "
    If m_Required Then
      jsString = jsString & "true"
    Else
      jsString = jsString & "false"
    End If
    For i = 0 To UBound(m_UploadRequest)
      Set Control = m_UploadRequest(i)
      jsString = jsString & ", ['" & Control.Name & "', '" & Control.AllowedExtensions & "', "
      If Control.Required Then
        jsString = jsString & "true]"
      Else
        jsString = jsString & "false]"
      End If
    Next
    jsString = jsString & ");showProgressWindow('" & m_ProgressTemplate & "?ProgressUrl=" & m_ScriptLibrary & "/PU3Progress.asp&UploadId=" & m_UploadId & "', " & m_ProgressWidth & ", " & m_ProgressHeight & ")"
    SubmitCode = jsString
  End Function
  
  Public Function ValidateCode()
    Dim jsString
    ' Generate JavaScript for filefield validation
    jsString = "validateFile(this, '" & m_AllowedExtensions & "', "
    If m_Required Then
      jsString = jsString & "true"
    Else
      jsString = jsString & "false"
    End If
    jsString = jsString & ")"
    ValidateCode = jsString
  End Function
  
  Private Function FieldExists(strName)
    Dim Index
    ' convert name to lowercase
    strName = LCase(strName)
      
    '  Loop through each field
    For Index = 0 To m_Count - 1
      '  If name matches current fields name in lowercase
      If LCase(m_UploadRequest(Index).Name) = strName Then
        '  Field Found
        FieldExists = True
        Exit Function
      End If
    Next
	FieldExists = False
  End Function

  Public Sub ProcessUpload()
    Dim Boundary, ChunkSize, BytesRead, reading, infoStream, dataStream, pos, myBuffer, BoundaryPos, DataPos, st, dt, real_fn, fn, s, i
    
    ' Check if form is posted
    If LCase(Request.ServerVariables("Request_Method")) <> "post" Then
      Exit Sub
    End If
    
    ' Check if ADO is available
    If Not Check_AdoDb() Then
      Exit Sub
    End If
    
    ' Check if form is sunmitted with correct contentType
    If Not Check_ContentType() Then
      Exit Sub
    End If
    
    If TotalBytes = 0 Then
      Response.Write "Nothing submitted"
      Response.End
    End If
    
    If m_UploadDone Then Exit Sub
    m_UploadDone = True

    If MaxSize > 0 Then
      If TotalBytes > MaxSize * 1024 Then
        On Error Resume Next
        Request.BinaryRead(TotalBytes) 'Need to read full Request
        If Err.Number <> 0 Then
          On Error GoTo 0
          Util.ThrowError 19, m_Translate.Value("PU3_ERR_BINARYREAD"), True
        Else
          On Error GoTo 0
          Util.ThrowError 6, m_Translate.ValueEx("PU3_ERR_FORMSIZE", Array(Util.ByteSize(TotalBytes), Util.ByteSize(m_MaxSize))), True
        End If
        Exit Sub
      End If
    End If

    Server.ScriptTimeout = m_TimeOut
	
    Set m_Progress = New UploadProgress
    m_Progress.RemoveAll
    m_Progress.UploadId = m_UploadId
    m_Progress.TimeOut = m_TimeOut
    m_Progress.Status = 0
    m_Progress.TotalBytes = TotalBytes
    m_Progress.UploadedBytes = 0
    Util.Progress = m_Progress
    
    Boundary = GetBoundary()
    ChunkSize = TotalBytes / 1024
    If ChunkSize < 1024 Then ChunkSize = 1024
    BytesRead = 0
    reading = "info"

    Dim Control, strName

    Set infoStream = Server.CreateObject("Adodb.Stream")
    infoStream.Mode = 3 'Read/Write
    infoStream.Type = 1 'Binary
    infoStream.Open()
    Set dataStream = Server.CreateObject("Adodb.Stream")
    dataStream.Mode = 3 'Read/Write
    dataStream.Type = 1 'Binary
    dataStream.Open()

    Do While ChunkSize > 0
      If ChunkSize + BytesRead > TotalBytes Then ChunkSize = TotalBytes - BytesRead
      If ChunkSize = 0 Then Exit Do
      
      m_Progress.Status = 1

      If reading = "info" Then
        On Error Resume Next
        infoStream.Write(Request.BinaryRead(ChunkSize))
        If Err.Number <> 0 Then
          On Error GoTo 0
          Util.ThrowError 19, m_Translate.Value("PU3_ERR_BINARYREAD"), True
        End If
        On Error GoTo 0
        infoStream.Position = 0
        myBuffer = infoStream.Read()
        pos = 0
      Else
        On Error Resume Next
        dataStream.Write(Request.BinaryRead(ChunkSize))
        If Err.Number <> 0 Then
          On Error GoTo 0
          Util.ThrowError 19, m_Translate.Value("PU3_ERR_BINARYREAD"), True
        End If
        On Error GoTo 0
        If dataStream.Position < (2 * ChunkSize) Then
          dataStream.Position = 0
          pos = 0
        Else
          pos = dataStream.Position - ChunkSize - LenB(Boundary)
          If pos < 0 Then pos = 0
          dataStream.Position = pos
        End If
        myBuffer = dataStream.Read()
      End If

      BoundaryPos = InstrB(myBuffer, Boundary)
      Do While BoundaryPos > 0
        If reading = "data" Then
          If Control.FileName = "" Then
            If Control.Required Then
              Util.ThrowErrorEx 40, m_Translate.Value("PU3_ERR_REQUIRED"), Control, m_HaltOnErrors
            End If
            copyStreamPart dataStream, infoStream, pos + (BoundaryPos - 3), False
            dataStream.Position = pos + (BoundaryPos - 3)
            dataStream.SetEOS
            Control.Size = dataStream.Position
            dataStream.Position = 0
            'If BoundaryPos > 3 Then
            '  Control.Value = GetUnicode(MidB(myBuffer, 1, BoundaryPos - 3))
            'End If
            Control.Value = GetUnicode(MidB(dataStream.Read, 1))
            ClearStream dataStream
            'copyStreamPart dataStream, infoStream, BoundaryPos - 1, True
          Else
            copyStreamPart dataStream, infoStream, pos + (BoundaryPos - 3), False
            dataStream.Position = pos + (BoundaryPos - 3)
            dataStream.SetEOS
            If Control.AllowedExtensions <> "" Then
              Dim ExtChk, ExtArr, FileExist
              ExtChk = True
              ExtArr = Split(Control.AllowedExtensions, ",")
              For i = 0 to UBound(ExtArr)
                If LCase(Trim(ExtArr(i))) = LCase(Control.Extension) Then
                  ExtChk = False
                End If
              Next
              If ExtChk Then
                Util.ThrowErrorEx 10, m_Translate.ValueEx("PU3_ERR_EXTENSION", Array(Control.AllowedExtensions)), Control, m_HaltOnErrors
              End If
            End If
            Control.Size = dataStream.Position
            If Control.MaxFileSize > 0 And Control.Size > (Control.MaxFileSize * 1024) Then
              Util.ThrowErrorEx 7, m_Translate.ValueEx("PU3_ERR_FILESIZE", Array(Control.FileName, Util.ByteSize(Control.Size), Util.ByteSize(Control.MaxFileSize*1024))), Control, m_HaltOnErrors
            End If
            If dataStream.Position > 0 Then
              dataStream.Position = 0
              If Control.Width = 0 Then
                s = GetSize(dataStream.Read(20000), Control.Extension)
                Control.Width = s(0)
                Control.Height = s(1)
              End If
              If Control.Width > 0 And Control.Height > 0 Then
                If Control.Width < Control.MinWidth Or  Control.Height < Control.MinHeight Then
                  Util.ThrowErrorEx 8, m_Translate.ValueEx("PU3_ERR_DIMSMALL", Array(Control.Width, Control.Height, Control.MinWidth, Control.MinHeight)), Control, m_HaltOnErrors
                ElseIf (Control.MaxWidth > 0 And Control.Width > Control.MaxWidth) Or (Control.MaxHeight > 0 And  Control.Height > Control.MaxHeight) Then
                  Util.ThrowErrorEx 9, m_Translate.ValueEx("PU3_ERR_DIMLARGE", Array(Control.Width, Control.Height, Control.MaxWidth, Control.MaxHeight)), Control, m_HaltOnErrors
                End If
              End If
              If Not m_KeepInMemory Then
                m_Progress.Status = 2
                m_Progress.SetFileInfo "Status", 2
                On Error Resume Next
                dataStream.SaveToFile Control.TempFileName, 2
                If Err.Number <> 0 Then
                  On Error GoTo 0
                  Dim filename, path
                  filename = Mid(Control.TempFileName, InStrRev(Control.TempFileName, "\")+1)
                  path = Left(Control.TempFileName, InStrRev(Control.TempFileName, "\"))
                  Util.ThrowError 21, m_Translate.ValueEx("PU3_ERR_SAVE", Array(filename, path)), True
                End If
                On Error GoTo 0
              End If
              If m_KeepInMemory Then
                Control.Blob = dataStream
              End If
              m_Progress.SetFileInfo "Status", 3
              ClearStream dataStream
            End If
          End If

          AddUploadFormRequest Control.Name, Control.FileName, Control.ContentType, Control.Value

          infoStream.Position = 0
          myBuffer = infoStream.Read()
          reading = "info"
        Else
          DataPos = InstrB(BoundaryPos, myBuffer, GetBinary(vbCrLf & vbCrLf))
          If DataPos > 0 Then
            st = BoundaryPos + LenB(Boundary)
            
            dt = GetUnicode(MidB(myBuffer, st, DataPos - st))
            
            strName = GetSubMatch(dt, " name=""", """")
            If Not FieldExists(strName) Then
              Set Control = New UploadControl
              Control.Translation = m_Translate
              Control.RaiseErrors = m_RaiseErrors
              Control.Name = strName
            Else
                Set Control = Fields(strName)
            End If

            real_fn = ""
            If InStr(dt, "filename=") > 0 Then
              fn = GetSubMatch(dt, " filename=""", """")
              real_fn = Mid(fn, InStrRev(fn, "\") + 1)
              real_fn = Mid(real_fn, InStrRev(real_fn, ":") + 1)
              m_Progress.AddFile real_fn
              m_Progress.SetFileInfo "Status", 1
              Control.FileName = real_fn
              Control.ContentType = GetSubMatch(dt, "Content-Type: ", vbCrLf)
              Control.Value = real_fn
              Control.StoreType = m_StoreType
              If Not m_KeepInMemory Then
                If m_TempFolder = "" Then m_TempFolder = UploadFolder
                m_TempFolder = Util.GetPhysicalPath(m_TempFolder)
                Util.AutoCreatePath m_TempFolder
                Control.TempFileName = m_TempFolder & "\" & GenerateTempFileName()
              End If
              If Control.UploadFolder = "" Then Control.UploadFolder = m_Path
              If Control.AllowedExtensions = "" Then Control.AllowedExtensions = m_AllowedExtensions
              If Control.ConflictHandling = puConflictOverwrite Then Control.ConflictHandling = m_ConflictHandling
              If Control.MinWidth = 0 Then Control.MinWidth = m_MinWidth
              If Control.MinHeight = 0 Then Control.MinHeight = m_MinHeight
              If Control.MaxWidth = 0 Then Control.MaxWidth = m_MaxWidth
              If Control.MaxHeight = 0 Then Control.MaxHeight = m_MaxHeight
              If Control.MaxFileSize = 0 Then Control.MaxFileSize = MaxFileSize
            End If

            If Not FieldExists(strName) Then
              ReDim Preserve m_UploadRequest(m_Count)
              Set m_UploadRequest(m_Count) = Control
              m_Count = m_Count + 1
            End If

            If DataPos + 3 < infoStream.Position Then
              copyStreamPart infoStream, dataStream, DataPos + 3, True
            End If
            dataStream.Position = 0
            myBuffer = dataStream.Read()
            pos = 0
            reading = "data"
          else
            'no header end in info buffer
            exit do
          End If
        End If

        BoundaryPos = InstrB(myBuffer, Boundary)

        If InstrB(myBuffer, Boundary & GetBinary("--")) = BoundaryPos And reading = "info" Then exit do

        If Not Response.IsClientConnected() Then
          If m_UseProgress Then
            DeleteProgressFile()
          End If
          infoStream.Close()
          dataStream.Close()
          Response.End
        End If
      Loop

      BytesRead = BytesRead + ChunkSize
      
      m_Progress.UploadedBytes = BytesRead

      If Not Response.IsClientConnected() Then
        infoStream.Close()
        dataStream.Close()
        Response.End
      End If
    Loop

    m_Progress.Status = 3

    infoStream.Close()
    dataStream.Close()
    
    ' Store querystring in formfield collection
    For Each strName In Request.QueryString
      If Not FieldExists(strName) Then
        Set Control = New UploadControl
        Control.Translation = m_Translate
        Control.Name = strName
        Control.Value = Request.QueryString(strName)
        If Not FieldExists(strName) Then
          ReDim Preserve m_UploadRequest(m_Count)
          Set m_UploadRequest(m_Count) = Control
          m_Count = m_Count + 1
          AddUploadFormRequest Control.Name, Control.FileName, Control.ContentType, Control.Value
        End If
      End If
    Next
  End Sub

  Public Sub SaveAll()
    Dim Index
    For Index = 0 To m_Count - 1
      m_UploadRequest(Index).Save
    Next
    pau_thePath = UploadFolder
  End Sub

  Private Function GenerateTempFileName()
    Dim fso : Set fso = Server.CreateObject("Scripting.FileSystemObject")
    GenerateTempFileName = fso.GetTempName
    Set fso = Nothing
  End Function
  
  Private Function Check_AdoDb()
    ' Verify ADODB Version
    Dim AdoDbConnection, AdoDbVersion
    Set AdoDbConnection = Server.CreateObject("AdoDb.Connection")
    AdoDbVersion = Replace(AdoDbConnection.Version, ".", "")
    AdoDbVersion = CInt(Left(AdoDbVersion, 2))
    Set AdoDbConnection = Nothing
    If AdoDbVersion < 25 Then
      Util.ThrowError 11, m_Translate.Value("PU3_ERR_ADO"), True
      Check_AdoDb = False
    Else
      Check_AdoDb = True
    End If
  End Function

  Private Function Check_ContentType()
    Dim ContentType
    ' Check content type
    ContentType = Request.ServerVariables("Http_Content_Type")
    If InStr(1, LCase(ContentType), "multipart/form-data") = 0 Then 
      'Util.ThrowError 17, m_Translate.Value("PU3_ERR_ENCTYPE"), True
      Check_ContentType = False
    Else
      Check_ContentType = True
    End If
  End Function

  Private Function GetBoundary()
    Dim ContentType, PosBeg, Boundary, Length
    ContentType = Request.ServerVariables("HTTP_Content_Type")

    ' Find the boundary
    PosBeg = InStr(1, ContentType, "boundary=")
    Boundary = getBinary(Mid(ContentType, PosBeg + 9))

    'bugfix IE5.01 - double header
    PosBeg = InStr(ContentType, "boundary=")
    If PosBeg > 0 Then
      PosBeg = InStr(Boundary, ",")
      If PosBeg > 0 Then Boundary = Left(Boundary, PosBeg - 1)
    End If

    Length = CLng(Request.ServerVariables("HTTP_Content_Length"))
    If Length > 0 And Boundary <> "" Then Boundary = getBinary("--") & Boundary

    GetBoundary = Boundary
  End Function

  Private Sub CopyStreamPart(ByRef fromStream, ByRef destStream, pos, clearAfterCopy)
    Dim tempStream
    Set tempStream = Server.CreateObject("Adodb.Stream")
    tempStream.Mode = 3 'adModeReadWrite
    tempStream.Type = 1 'adTypeBinary
    tempStream.Open()

    If pos < fromStream.Position Then fromStream.Position = pos
    fromStream.CopyTo tempStream
    If clearAfterCopy Then clearStream fromStream

    tempStream.Position = 0
    tempStream.CopyTo destStream
    tempStream.Position = 0
    tempStream.SetEOS

    tempStream.Close()
  End Sub

  Private Sub ClearStream(ByRef stream)
    stream.Position = 0
    stream.SetEOS
  End Sub

  Private Function GetSubMatch(str, sStart, sEnd)
    Dim posStart, posEnd
    posStart = InStr(str, sStart)
    posStart = posStart + Len(sStart)
    posEnd = InStr(posStart, str, sEnd)
    If posStart > 0 And posEnd > posStart Then
      GetSubMatch = Mid(str, posStart, posEnd - posStart)
    Else
      GetSubMatch = ""
    End If
  End Function

  Private Function GetBinary(strUnicode)
    Dim Index
    For Index = 1 To Len(strUnicode)
      GetBinary = GetBinary & ChrB(Asc(Mid(strUnicode, Index, 1)))
    Next
  End Function

  Private Function GetUnicodeSimple(str)
    Dim Index
    For Index = 1 To LenB(str)
      GetUnicodeSimple = GetUnicodeSimple & Chr(AscB(MidB(str, Index, 1)))
    Next
  End Function
  
  Function GetUnicode(str)
    If m_CharSet <> "" Then
      Dim BinaryStream, outStr
      Set BinaryStream = Server.CreateObject("ADODB.Stream")
      BinaryStream.Type = 2
      BinaryStream.Open
      str = GetBinary("....DMX") & str
      BinaryStream.WriteText str
      BinaryStream.Position = 0
      BinaryStream.CharSet = m_CharSet
      outStr = BinaryStream.ReadText
      GetUnicode = Mid(outStr, InStr(outStr, "DMX")+3)
    Else
      GetUnicode = GetUnicodeSimple(str)
    End If
  End Function
  
  Private Sub initCharsetMap()
    m_CharSetMap = Array("Arabic (ASMO 708)","ASMO-708","708", _
          "Arabic (ISO)","iso-8859-6","28596", _
          "Arabic (Windows)","windows-1256","1256", _
          "Baltic (ISO)","iso-8859-4","28594", _
          "Baltic (Windows)","windows-1257","1257", _
          "Central European (DOS)","ibm852","852", _
          "Central European (ISO)","iso-8859-2","28592", _
          "Central European (Windows)","windows-1250","1250", _
          "Chinese Simplified (EUC)","EUC-CN","51936", _
          "Chinese Simplified (GB2312)","gb2312","936", _
          "Chinese Simplified (HZ)","hz-gb-2312","52936", _
          "Chinese Traditional (Big5)","big5","950", _
          "Cyrillic (DOS)","cp866","866", _
          "Cyrillic (ISO)","iso-8859-5","28595", _
          "Cyrillic (KOI8-R)","koi8-r","20866", _
          "Cyrillic (KOI8-U)","koi8-u","21866", _
          "Cyrillic (Windows)","windows-1251","1251", _
          "Greek (ISO)","iso-8859-7","28597", _
          "Greek (Windows)","windows-1253","1253", _
          "Hebrew (DOS)","DOS-862","862", _
          "Hebrew (ISO-Logical)","iso-8859-8-i","38598", _
          "Hebrew (ISO-Visual)","iso-8859-8","28598", _
          "Hebrew (Windows)","windows-1255","1255", _
          "Japanese (EUC)","euc-jp","51932", _
          "Japanese (JIS)","iso-2022-jp","50220", _
          "Japanese (Shift-JIS)","shift_jis","932", _
          "Korean (EUC)","euc-kr","51949", _
          "Korean (ISO)","iso-2022-kr","50225", _
          "Korean (Johab)","Johab","1361", _
          "Latin 3 (ISO)","iso-8859-3","28593", _
          "Latin 9 (ISO)","iso-8859-15","28605", _
          "Norwegian (IA5)","x-IA5-Norwegian","20108", _
          "Swedish (IA5)","x-IA5-Swedish","20107", _
          "Thai (Windows)","windows-874","874", _
          "Turkish (ISO)","iso-8859-9","28599", _
          "Turkish (Windows)","windows-1254","1254", _
          "Unicode (UTF-8)","utf-8","65001", _
          "US-ASCII","us-ascii","20127", _
          "Vietnamese (Windows)","windows-1258","1258", _
          "Western European (ISO)","iso-8859-1","28591", _
          "Western European (Windows)","Windows-1252","1252" )
  End Sub

  Private Function GetSize(buff, extension)
    Dim w, h, i, lngBuffer, bBuf(), lngPos, nBits, RECTdata, xMin, xMax, yMin, yMax

    lngBuffer = LenB(buff)
    ReDim bBuf(lngBuffer)
    For i = 1 To lngBuffer
      bBuf(i-1) = AscB(MidB(buff, i, 1))
    Next
    
    Select Case LCase(extension)
      Case "png"
        If bBuf(0) = 137 And bBuf(1) = 80 And bBuf(2) = 78 Then
          w = Mult(bBuf(19), bBuf(18))
          h = Mult(bBuf(23), bBuf(22))
          GetSize = Array(w,h)
          Exit Function
        End If
      Case "gif"
        If bBuf(0) = 71 And bBuf(1) = 73 And bBuf(2) = 70 Then
          w = Mult(bBuf(6), bBuf(7))
          h = Mult(bBuf(8), bBuf(9))
          GetSize = Array(w,h)
          Exit Function
        End If
      Case "bmp"
        If bBuf(0) = 66 And bBuf(1) = 77 Then
          w = Mult(bBuf(18), bBuf(19))
          h = Mult(bBuf(22), bBuf(23))
          GetSize = Array(w,h)
          Exit Function
        End If
      Case "swf" 'only uncompressed
        If bBuf(0) = 70 And bBuf(1) = 87 And bBuf(2) = 83 Then 'bBuf(0) = 67 compressed
          RECTdata = ToBin(bBuf(8), 8)
          RECTdata = RECTdata & ToBin(bBuf(9), 8)
          RECTdata = RECTdata & ToBin(bBuf(10), 8)
          RECTdata = RECTdata & ToBin(bBuf(11), 8)
          RECTdata = RECTdata & ToBin(bBuf(12), 8)
          RECTdata = RECTdata & ToBin(bBuf(13), 8)
          RECTdata = RECTdata & ToBin(bBuf(14), 8)
          RECTdata = RECTdata & ToBin(bBuf(15), 8)
          RECTdata = RECTdata & ToBin(bBuf(16), 8)
  
          nBits = Mid(RECTdata, 1, 5)
          nBits = Bin2Decimal(nBits)
  
          xMin = Bin2Decimal(Mid(RECTdata, 6, nBits))
          xMax = Bin2Decimal(Mid(RECTdata, 6 + (nBits * 1), nBits))
          yMin = Bin2Decimal(Mid(RECTdata, 6 + (nBits * 2), nBits))
          yMax = Bin2Decimal(Mid(RECTdata, 6 + (nBits * 3), nBits))
  
          w = (xMax - xMin) / 20
          h = (yMax - yMin) / 20
          GetSize = Array(w,h)
          Exit Function
        End If
      Case "avi"
        If bBuf(0) = 82 And bBuf(1) = 73 And bBuf(2) = 70 And bBuf(3) = 70 And bBuf(8) = 65 And bBuf(9) = 86 And bBuf(10) = 73 Then
          If bBuf(24) = 97 And bBuf(25) = 118 And bBuf(26) = 105 And bBuf(27) = 104 Then
            w = Mult(bBuf(64), bBuf(65))
            h = Mult(bBuf(68), bBuf(69))
            GetSize = Array(w,h)
            Exit Function
          End If
        End If
      Case "mov" 'only uncompressed
        If bBuf(4) = 109 And bBuf(5) = 111 And bBuf(6) = 111 And bBuf(7) = 118 Then
          If bBuf(12) = 109 And bBuf(13) = 118 And bBuf(14) = 104 And bBuf(15) = 100 Then
            If bBuf(128) = 116 And bBuf(129) = 107 And bBuf(130) = 104 And bBuf(131) = 100 Then
              w = (bBuf(208) * CLng(256)) + bBuf(209)
              h = (bBuf(212) * CLng(256)) + bBuf(213)
              GetSize = Array(w,h)
              Exit Function
            End If
          End If
        End If
      Case "tga"
        w = Mult(bBuf(12), bBuf(13))
        h = Mult(bBuf(14), bBuf(15))
        GetSize = Array(w,h)
        Exit Function
      Case "tif", "tiff"
        If ((bBuf(0) = 73 And bBuf(1) = 73) Or (bBuf(0) = 77 And bBuf(1) = 77)) And bBuf(2) = 42 Then
          lngPos = Mult32(bBuf(4), bBuf(5), bBuf(6), bBuf(7))

          w = 0
          h = 0

          For i = 1 To Mult(bBuf(lngPos), bBuf(lngPos + 1))
            If bBuf(lngPos + 2) = &H00 And bBuf(lngPos + 3) = &H01 Then
              w = Mult(bBuf(lngPos + 10), bBuf(lngPos + 11))
            End If

            If bBuf(lngPos + 2) = &H01 And bBuf(lngPos + 3) = &H01 Then
              h = Mult(bBuf(lngPos + 10), bBuf(lngPos + 11))
              GetSize = Array(w,h)
              Exit Function
            End If

            lngPos = lngPos + 12
          Next
        End If
      Case "jpg", "jpeg"
        If bBuf(0) = 255 And bBuf(1) = 216 And bBuf(2) = 255 Then
          Do
            If (bBuf(lngPos) = &HFF And bBuf(lngPos + 1) = &HD8 And bBuf(lngPos + 2) = &HFF) Or (lngPos >= lngBuffer - 10) Then Exit Do
            lngPos = lngPos + 1
          Loop

          lngPos = lngPos + 2
          If lngPos >= lngBuffer - 10 Then
            w = 0
            h = 0
            GetSize = Array(w,h)
            Exit Function
          End If

          Do
            Do
              If bBuf(lngPos) = &HFF And bBuf(lngPos + 1) <> &HFF Then Exit Do
              lngPos = lngPos + 1
              If lngPos >= lngBuffer - 10 Then
                w = 0
                h = 0
                GetSize = Array(w,h)
                Exit Function
              End If
            Loop

            lngPos = lngPos + 1

            If bBuf(lngPos) >= &HC0 And bBuf(lngPos) <= &HC3 Then Exit Do
            If bBuf(lngPos) >= &HC5 And bBuf(lngPos) <= &HC7 Then Exit Do
            If bBuf(lngPos) >= &HC9 And bBuf(lngPos) <= &HCB Then Exit Do
            If bBuf(lngPos) >= &HCD And bBuf(lngPos) <= &HCF Then Exit Do

          lngPos = lngPos + Mult(bBuf(lngPos + 2), bBuf(lngPos + 1))

            If lngPos >= lngBuffer - 10 Then
              w = 0
              h = 0
              GetSize = Array(w,h)
              Exit Function
            End If
          Loop

          w = Mult(bBuf(lngPos + 7), bBuf(lngPos + 6))
          h = Mult(bBuf(lngPos + 5), bBuf(lngPos + 4))
          GetSize = Array(w,h)
          Exit Function
        End If
    End Select

    w = 0
    h = 0
    GetSize = Array(w,h)
  End Function

  Private Function Mult(lsb, msb)
    Mult = lsb + (msb * CLng(256))
  End Function
  
  Private Function Mult32(b1, b2, b3, b4)
    Mult32 = b1 + (b2 * CLng(256)) + (b3 * CLng(65536)) + (b4 * CLng(16777216))
  End Function

  Private Function ToBin(inNumber, OutLenStr)
    Dim binary
    binary = ""
    Do While inNumber >= 1
      binary = binary & inNumber Mod 2
      inNumber = inNumber \ 2
    Loop
    binary = binary & String(OutLenStr - len(binary), "0")
    ToBin = StrReverse(binary)
  End Function

  Private Function Bin2Decimal(inBin)
    Dim counter, temp, Value
    inBin = StrReverse(inBin)
    temp = 0
    For counter = 1 To Len(inBin)
      If counter = 1 Then
        Value = 1
      Else
        Value = Value * 2
      End If
      temp = temp + Mid(inBin, counter, 1) * Value
    Next
    Bin2Decimal = temp
  End Function
End Class


Class UploadControl

  Private m_Name
  Private m_FileName
  Private m_BaseName
  Private m_Extension
  Private m_TempFileName
  Private m_Path
  Private m_ContentType
  Private m_Value
  Private m_Size
  Private m_Width
  Private m_Height
  Private m_MaxFileSize
  Private m_AllowedExtensions
  Private m_MinWidth
  Private m_MinHeight
  Private m_MaxWidth
  Private m_MaxHeight
  Private m_ConflictHandling
  Private m_Required
  Private m_Blob
  Private m_Error
  Private m_Translate
  Private m_StoreType
  
  Private Util

  Public Property Let RaiseErrors(bool)
    Util.RaiseErrors = bool
  End Property

  Public Property Let Translation(obj)
    Set m_Translate = obj
    Util.Translation = obj
  End Property

  Public Property Let StoreType(strOption)
    m_StoreType = Trim(CStr(strOption))
    If m_StoreType <> puStoreFile And m_StoreType <> puStorePath Then
      Util.ThrowError 4, m_Translate.ValueEx("PU3_ERR_VALUE", Array("StoreType")), True
    End If
  End Property
  
  Public Property Get StoreType()
    StoreType = m_StoreType
  End Property
  
  Public Property Let Name(strName)
    m_Name = CStr(strName)
  End Property

  Public Property Get Name()
    Name = m_Name
  End Property

  Public Property Let FileName(strFileName)
    m_FileName = Trim(CStr(strFileName))
    If InStr(m_FileName, "\") > 0 Then
      m_FileName = Mid(m_FileName, InStrRev(m_FileName, "\") + 1)
    ElseIf InStr(m_FileName, "/") > 0 Then
      m_FileName = Mid(m_FileName, InStrRev(m_FileName, "/") + 1)
    End If
    m_BaseName = Mid(m_FileName, 1, Len(m_FileName) - InStrRev(m_FileName, "."))
    m_Extension = Mid(m_FileName, InStrRev(m_FileName, ".") + 1)
  End Property

  Public Property Get FileName()
    FileName =m_FileName
  End Property

  Public Property Get BaseName()
    BaseName = m_BaseName
  End Property

  Public Property Get Extension()
    Extension = m_Extension
  End Property

  Public Property Let TempFileName(strTempFileName)
    m_TempFileName = Trim(CStr(strTempFileName))
  End Property

  Public Property Get TempFileName()
    TempFileName = m_TempFileName
  End Property

  Public Property Let UploadFolder(strPath)
    m_Path = Trim(CStr(strPath))
  End Property

  Public Property Get UploadFolder()
    If InStr(m_Path, """") > 0 Then
      UploadFolder = Eval(m_Path)
    Else
      UploadFolder = m_Path
    End If
  End Property

  Public Property Let ContentType(strContentType)
    m_ContentType = CStr(strContentType)
  End Property

  Public Property Get ContentType()
    ContentType = m_ContentType
  End Property

  Public Property Let Value(strValue)
    m_Value = CStr(strValue)
  End Property

  Public Default Property Get Value()
    Value = m_Value
  End Property

  Public Property Let Size(lngSize)
    m_Size = CLng(lngSize)
  End Property

  Public Property Get Size()
    Size = m_Size
  End Property

  Public Property Let Width(lngSize)
    m_Width = CLng(lngSize)
  End Property

  Public Property Get Width()
    Width = m_Width
  End Property

  Public Property Let Height(lngSize)
    m_Height = CLng(lngSize)
  End Property

  Public Property Get Height()
    Height = m_Height
  End Property
  
  Public Property Let MaxFileSize(lngSize)
    m_MaxFileSize = CLng(lngSize*1024)
  End Property
  
  Public Property Get MaxFileSize()
    MaxFileSize = m_MaxFileSize/1024
  End Property

  Public Property Let AllowedExtensions(strExtensions)
    m_AllowedExtensions = Trim(CStr(strExtensions))
  End Property  

  Public Property Get AllowedExtensions()
    AllowedExtensions = m_AllowedExtensions
  End Property  

  Public Property Let MinWidth(intWidth)
    m_MinWidth = CLng(Abs(intWidth))
  End Property
  
  Public Property Get MinWidth()
    MinWidth = m_MinWidth
  End Property
  
  Public Property Let MinHeight(intHeight)
    m_MinHeight = CLng(Abs(intHeight))
  End Property
  
  Public Property Get MinHeight()
    MinHeight = m_MinHeight
  End Property
  
  Public Property Let MaxWidth(intWidth)
    m_MaxWidth = CLng(Abs(intWidth))
  End Property
  
  Public Property Get MaxWidth()
    MaxWidth = m_MaxWidth
  End Property
  
  Public Property Let MaxHeight(intHeight)
    m_MaxHeight = CLng(Abs(intHeight))
  End Property
  
  Public Property Get MaxHeight()
    MaxHeight = m_MaxHeight
  End Property
  
  Public Property Let ConflictHandling(strOption)
    m_ConflictHandling = Trim(CStr(strOption))
    If m_ConflictHandling <> puConflictIgnore And m_ConflictHandling <> puConflictUnique And m_ConflictHandling <> puConflictError And m_ConflictHandling <> puConflictOverwrite Then
      Util.ThrowError 4, m_Translate.ValueEx("PU3_ERR_VALUE", Array("ConflictHandling")), True
    End If
  End Property

  Public Property Get ConflictHandling()
    ConflictHandling = m_ConflictHandling
  End Property
  
  Public Property Let Required(boolReq)
    m_Required = CBool(boolReq)
  End Property
  
  Public Property Get Required()
    Required = m_Required
  End Property

  Public Property Let Blob(objBlob)
    Set m_Blob = Server.CreateObject("Adodb.Stream")
    m_Blob.Mode = 3 'adModeReadWrite
    m_Blob.Type = 1 'adTypeBinary
    m_Blob.Open()

    objBlob.Position = 0
    objBlob.CopyTo m_Blob
  End Property

  Public Property Get Blob()
    Dim binBlob
    binBlob = ""
    If IsObject(m_Blob) Then
      m_Blob.Position = 0
      binBlob = m_Blob.Read()
    End If
    Blob = binBlob
  End Property
  
  Public Property Let Error(strError)
    m_Error = CStr(strError)
  End Property
  
  Public Property Get Error()
    Error = m_Error
  End Property

  Private Sub Class_Initialize()
    m_Name = ""
    m_FileName = ""
    m_BaseName = ""
    m_Extension = ""
    m_TempFileName = ""
    m_Path = ""
    m_ContentType = ""
    m_Value = ""
    m_Size = 0
    m_Width = 0
    m_Height = 0
    m_MaxFileSize = 0
    m_AllowedExtensions = ""
    m_MinWidth = 0
    m_MinHeight = 0
    m_MaxWidth = 0
    m_MaxHeight = 0
    m_ConflictHandling = puConflictOverwrite
    m_Required = False
    Set Util = New PureUploadUtils
  End Sub

  Private Sub Class_Terminate()
    Set m_Translate = Nothing
    Set Util = Nothing
    If IsObject(m_Blob) Then
      m_Blob.Close
    End If
  End Sub
  
  Public Function ValidateCode()
    Dim jsString
    ' Generate JavaScript for filefield validation
    jsString = "validateFile(this, '" & m_AllowedExtensions & "', "
    If m_Required Then
      jsString = jsString & "true"
    Else
      jsString = jsString & "false"
    End If
    jsString = jsString & ")"
    ValidateCode = jsString
  End Function
  
  Public Sub Save()
    Dim curPath, pos, absPath, Fso, i, FileExist
    If m_FileName <> "" And m_Error = "" Then
      'First do some checks
      If m_Size = 0 Then
        Util.ThrowErrorEx 20, m_Translate.ValueEx("PU3_ERR_EMPTY", Array(m_FileName)), Me, m_HaltOnErrors
        Exit Sub
      End If
      
      absPath = Util.GetPhysicalPath(UploadFolder)
      Util.AutoCreatePath absPath
      absPath = absPath & "\"

      Set Fso = Server.CreateObject("Scripting.FileSystemObject")
      FileExist = Fso.FileExists(absPath & FileName)

      If m_ConflictHandling = puConflictError And FileExist Then
        Set Fso = Nothing
        Util.ThrowError 12, m_Translate.ValueEx("PU3_ERR_EXISTS", Array(m_FileName)), True
        Exit Sub
      End If

      If ((m_ConflictHandling = puConflictOverwrite Or m_ConflictHandling = puConflictUnique) And FileExist) or (Not FileExist) Then
        If m_ConflictHandling = puConflictUnique And FileExist Then
          Dim num, newFileName
          num = 0
          newFileName = absPath & m_FileName
          While FileExist
            num = num + 1
            newFileName = Fso.GetBaseName(absPath & m_FileName) & "_" & num & "." & Fso.GetExtensionName(absPath & m_FileName)
            FileExist = Fso.FileExists(absPath & newFileName)
          Wend
          m_FileName = newFileName
          SetUploadFormRequestFileName m_Name, m_FileName, m_ContentType
		  SetUploadFormRequestValue m_Name, m_FileName
        End If
        
        If m_StoreType = puStorePath And UploadFolder <> "" Then
          If InStr(UploadFolder, "\") > 0 Then
            If Right(UploadFolder, 1) = "\" Then
              m_Value = UploadFolder & m_FileName
            Else
              m_Value = UploadFolder & "\" & m_FileName
            End If
          Else
            If Left(UploadFolder, 1) = "/" Then
              curPath = UploadFolder
            Else
              curPath = Request.ServerVariables("PATH_INFO")
              curPath = Left(curPath, InStrRev(curPath, "/")) & UploadFolder
              curPath = Replace(curPath, "/./", "/")
              While InStr(curPath, "/../") > 0
                pos = InStr(curPath, "/../")
                If pos > 1 Then
                  curPath = Left(curPath, InStrRev(curPath, "/", pos-1)) & Mid(curPath, pos+4)
                Else
                  curPath = Left(curPath, pos) & Mid(curPath, pos+4)
                End If
              Wend
            End If
            If Right(curPath, 1) <> "/" Then curPath = curPath & "/"
            m_Value = curPath & m_FileName
          End If
          SetUploadFormRequestValue m_Name, m_Value
        End If
      
        If m_TempFileName <> "" Then
          If Fso.FileExists(absPath & m_FileName) Then Fso.DeleteFile(absPath & m_FileName)
          Fso.MoveFile m_TempFileName, absPath & m_FileName
        ElseIf IsObject(m_Blob) Then
          m_Blob.SaveToFile absPath & m_FileName, 2
        End If
      End If

      Set Fso = Nothing
    End If
  End Sub
  
End Class


Class PureUploadUtils
  
  Private m_Translate
  Private m_Progress
  Private m_RaiseErrors
  
  Public Property Let Translation(obj)
    Set m_Translate = obj
  End Property
  
  Public Property Let Progress(obj)
    Set m_Progress = obj
  End Property
  
  Public Property Let RaiseErrors(bool)
    m_RaiseErrors = bool
  End Property
  
  Private Sub Class_Initialize()
    m_RaiseErrors = False
  End Sub
  
  Private Sub Class_Terminate()
    Set m_Translate = Nothing
    Set m_Progress = Nothing
  End Sub
  
  Public Function ByteSize(lngBytes)
    If lngBytes < 1024 Then
      ByteSize = FormatNumber(lngBytes, 0) & "B"
      Exit Function
    ElseIf lngBytes < 1048576 Then
      ByteSize = FormatNumber(lngBytes/1024, 0) & "kB"
      Exit Function
    ElseIf lngBytes < 1073741824 Then
      ByteSize = FormatNumber(lngBytes/1048576, 2) & "MB"
      Exit Function
    Else
      ByteSize = FormatNumber(lngBytes/1073741824, 2) & "GB"
      Exit Function
    End If
  End Function

  Public Function GetPhysicalPath(strPath)
    Dim curPath, arrPath, newPath, backCount, i
    newPath = strPath
    If newPath = "" Then newPath = "."
    If InStr(newPath, "\") = 0 Then
      If InStr(newPath, "../") = 0 Then
        newPath = Server.MapPath(newPath)
      Else
        curPath = Server.MapPath(".")
        If newPath = "" Then
          GetPhysicalPath = curPath
          Exit Function
        End If
        arrPath = Split(curPath & "\" & Replace(newPath, "/", "\"), "\")
        newPath = ""
        backCount = 0
        For i = UBound(arrPath) To 1 Step -1
          If arrPath(i) = ".." Then
            backCount = backCount + 1
          Else
            If backCount = 0 Then
              newPath = arrPath(i) & "\" & newPath
            Else
              backCount = backCount - 1
            End If
          End If
        Next
        If backCount > 0 Then
          ' Incorrect Path
          ThrowError 15, m_Translate.ValueEx("PU3_ERR_FOLDER", Array(strPath)), True
        End If
        newPath = arrPath(0) & "\" & Left(newPath, Len(newPath) - 1)
      End If
    End If
    If Right(newPath, 1) = "\" Then newPath = Left(newPath, Len(newPath) - 1)
    GetPhysicalPath = newPath
  End Function
  
  Public Sub AutoCreatePath(strPath)
    Dim Fso, Pos, NewPath

    Set Fso = Server.CreateObject("Scripting.FileSystemObject")
    If Not Fso.FolderExists(strPath) Then
      Pos = InStrRev(strPath, "\")
      If Pos > 0 Then
        NewPath = Left(strPath, Pos - 1)
        AutoCreatePath NewPath
        On Error Resume Next
        Fso.CreateFolder strPath
        If Err.Number <> 0 Then
          On Error Goto 0
          ThrowError 13, m_Translate.ValueEx("PU3_ERR_CREATEFOLDER", Array(strPath)), True
        End If
        On Error Goto 0
      End If  
    End If
    Set Fso = Nothing
  End Sub
  
  Sub ThrowError(nr, str, raise)
    If IsObject(m_Progress) Then
      m_Progress.AddGlobalError str
    End If
    
    If raise Then
      If m_RaiseErrors Then
        Err.Raise nr, "PureUpload", str
      Else
        Response.Write "<h1>Upload Error</h1>"
        Response.Write "(" & nr & ") " & str
        Response.End
      End If
    End If
  End Sub
  
  Sub ThrowErrorEx(nr, str, Control, raise)
    If IsObject(m_Progress) Then
      m_Progress.SetFileInfo "Error", str
    End If
    
    Control.Error = str
    
    If raise Then
      If m_RaiseErrors Then
        Err.Raise nr, "PureUpload", str
      Else
        Response.Write "<h1>Upload Error</h1>"
        Response.Write "(" & nr & ") " & str
        Response.End
      End If
    End If
  End Sub
  
End Class

'------------------------------------------
' Make the class accessable by JScript
'------------------------------------------

Function CreatePureUpload()
  Set CreatePureUpload = New PureUpload
End Function

'------------------------------------------
' Functions for backwards compatibility with version 2
'------------------------------------------

' Fill the UploadRequest
Sub AddUploadFormRequest(strName, strFileName, strContentType, strValue)
  strName = LCase(strName)
  If UploadRequest.Exists(strName) then
    UploadRequest(strName).Item("Value") = UploadRequest(strName).Item("Value") & "," & strValue
  Else
    Dim d : Set d = Server.CreateObject("Scripting.Dictionary")
    d.Add "FileName", strFileName
    d.Add "ContentType", strContentType
    d.Add "Value", strValue
    UploadRequest.Add strName, d 
    Set d = Nothing
  End If    
End Sub

Sub SetUploadFormRequest(strName, strValue)
  strName = LCase(strName)
  If UploadRequest.Exists(strName) Then
    UploadRequest(strName).Item("Value") = strValue
  Else
    Dim d : Set d = Server.CreateObject("Scripting.Dictionary")
    d.Add "FileName", ""
    d.Add "ContentType", ""
    d.Add "Value", strValue
    UploadRequest.Add strName, d 
    Set d = Nothing
  End If    
End Sub

Sub SetUploadFormRequestValue(strName, strValue)
  strName = LCase(strName)
  If UploadRequest.Exists(strName) Then
    UploadRequest(strName).Item("Value") = strValue
  Else
    Dim d : Set d = Server.CreateObject("Scripting.Dictionary")
    d.Add "FileName", ""
    d.Add "ContentType", ""
    d.Add "Value", strValue
    UploadRequest.Add strName, d 
    Set d = Nothing
  End If    
End Sub

Sub SetUploadFormRequestFileName(strName, strFileName, strContentType)
  strName = LCase(strName)
  If UploadRequest.Exists(strName) Then
    UploadRequest(strName).Item("FileName") = strFileName
    UploadRequest(strName).Item("ContentType") = strContentType
  Else
    Dim d : Set d = Server.CreateObject("Scripting.Dictionary")
    d.Add "FileName", strFileName
    d.Add "ContentType", strContentType
    d.Add "Value", strFileName
    UploadRequest.Add strName, d 
    Set d = Nothing
  End If    
End Sub

' Replacement for the requests
Function UploadFormRequest(strName)
  UploadFormRequest = ""
  strName = LCase(strName)
  If IsObject(UploadRequest) Then
    If UploadRequest.Exists(strName) Then
      If UploadRequest.Item(strName).Exists("Value") Then
        UploadFormRequest = UploadRequest.Item(strName).Item("Value")
      End If 
    End If  
  End If  
End Function

'Fix for the update record
Function FixFieldsForUpload(GP_fieldsStr, GP_columnsStr)
  Dim GP_counter, GP_Fields, GP_Columns, GP_FieldName, GP_FieldValue, GP_CurFileName, GP_CurContentType
  GP_Fields = Split(GP_fieldsStr, "|")
  GP_Columns = Split(GP_columnsStr, "|") 
  GP_fieldsStr = ""
  ' Get the form values
  For GP_counter = LBound(GP_Fields) To UBound(GP_Fields) Step 2
    GP_FieldName = LCase(GP_Fields(GP_counter))
    GP_FieldValue = GP_Fields(GP_counter+1)
    If UploadRequest.Exists(GP_FieldName) Then
      GP_CurFileName = UploadRequest.Item(GP_FieldName).Item("FileName")
      GP_CurContentType = UploadRequest.Item(GP_FieldName).Item("ContentType")
    Else  
      GP_CurFileName = ""
      GP_CurContentType = ""
    End If  
    If (GP_CurFileName = "" And GP_CurContentType = "") Or (GP_CurFileName <> "" And GP_CurContentType <> "") Then
      GP_fieldsStr = GP_fieldsStr & GP_FieldName & "|" & GP_FieldValue & "|"
    End If 
  Next
  If GP_fieldsStr <> "" Then
    GP_fieldsStr = Mid(GP_fieldsStr,1,Len(GP_fieldsStr)-1)
  Else  
    Response.Write "<strong>An error has occured during record update!</strong><br/><br/>"
    Response.Write "There are no fields to update ...<br/>"
    Response.Write "If the file upload field is the only field on your form, you should make it required.<br/>"
    Response.Write "Please correct and <a href=""javascript:history.back(1)"">try again</a>"
    Response.End
  End If
  FixFieldsForUpload = GP_fieldsStr    
End Function

'Fix for the update record
Function FixColumnsForUpload(GP_fieldsStr, GP_columnsStr)
  Dim GP_counter, GP_Fields, GP_Columns, GP_FieldName, GP_ColumnName, GP_ColumnValue,GP_CurFileName, GP_CurContentType
  GP_Fields = Split(GP_fieldsStr, "|")
  GP_Columns = Split(GP_columnsStr, "|") 
  GP_columnsStr = "" 
  ' Get the form values
  For GP_counter = LBound(GP_Fields) To UBound(GP_Fields) Step 2
    GP_FieldName = LCase(GP_Fields(GP_counter))  
    GP_ColumnName = GP_Columns(GP_counter)  
    GP_ColumnValue = GP_Columns(GP_counter+1)
    If UploadRequest.Exists(GP_FieldName) Then
      GP_CurFileName = UploadRequest.Item(GP_FieldName).Item("FileName")
      GP_CurContentType = UploadRequest.Item(GP_FieldName).Item("ContentType")    
    Else  
      GP_CurFileName = ""
      GP_CurContentType = ""
    End If  
    If (GP_CurFileName = "" And GP_CurContentType = "") Or (GP_CurFileName <> "" And GP_CurContentType <> "") Then
      GP_columnsStr = GP_columnsStr & GP_ColumnName & "|" & GP_ColumnValue & "|"
    End If 
  Next
  If GP_columnsStr <> "" Then
    GP_columnsStr = Mid(GP_columnsStr,1,Len(GP_columnsStr)-1)    
  End If
  FixColumnsForUpload = GP_columnsStr
End Function

</SCRIPT>