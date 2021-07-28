<SCRIPT LANGUAGE="VBSCRIPT" RUNAT="SERVER">
'*** Pure ASP File Upload -----------------------------------------------------
' Copyright 2001-2006 (c) George Petrov, www.DMXzone.com
' Version: 2.21
'------------------------------------------------------------------------------

'Current version
Function getPureUploadVersion()
  getPureUploadVersion = 2.21
End Function


'Set the querystring correctly
Sub PureUploadSetup()
  If (CStr(Request.QueryString("GP_upload")) <> "") Then
    UploadQueryString = Replace(Request.QueryString,"GP_upload=true","")
    if left(UploadQueryString,1) = "&" or left(UploadQueryString,1) = "?" then
      UploadQueryString = Mid(UploadQueryString,2)
    end if
    if right(UploadQueryString,1) = "&" then
      UploadQueryString = Mid(UploadQueryString,1,len(UploadQueryString)-1)
    end if  
  else  
    UploadQueryString = Request.QueryString
    If (UploadQueryString <> "") Then  
      UploadQueryString = UploadQueryString & "&GP_upload=true"
    else
      UploadQueryString = "GP_upload=true"
    end if  
    GP_uploadAction = CStr(Request.ServerVariables("URL")) & "?" & UploadQueryString
  end if  
End Sub


'Read the form(actual upload)
Sub ProcessUpload(pau_thePath,pau_Extensions,pau_Redirect,pau_storeType,pau_sizeLimit,pau_nameConflict,pau_requireUpload,pau_minWidth,pau_minHeight,pau_maxWidth,pau_maxHeight,pau_saveWidth,pau_saveHeight,pau_timeout)
  Server.ScriptTimeout = pau_timeout
  pau_doPreUploadChecks pau_sizeLimit
  RequestBin = Request.BinaryRead(Request.TotalBytes)
  Set UploadRequest = CreateObject("Scripting.Dictionary")  
  pau_BuildUploadRequest RequestBin, pau_thePath, pau_storeType, pau_sizeLimit, pau_nameConflict, pau_Extensions
  If pau_Redirect <> "" Then
    If UploadQueryString <> "" Then
      pau_Redirect = pau_Redirect & "?" & UploadQueryString
    End If
    Response.Redirect(pau_Redirect)  
  end if  
End Sub


'Some checks before actual upload
Sub pau_doPreUploadChecks(sizeLimit)
  Dim checkADOConn, AdoVersion, Length
  'Check ADO Version
  set checkADOConn = Server.CreateObject("ADODB.Connection")
  on error resume next
  adoVersion = CSng(checkADOConn.Version)
  if err then 
    adoVersion = Replace(checkADOConn.Version,".",",")  
    adoVersion = CSng(adoVersion)
  end if  
  err.clear
  on error goto 0 
  set checkADOConn = Nothing
  if adoVersion < 2.5 then
    Response.Write "<strong>You don't have ADO 2.5 installed on the server.</strong><br/>"
    Response.Write "The File Upload extension needs ADO 2.5 or greater to run properly.<br/>"
    Response.Write "You can download the latest MDAC (ADO is included) from <a href=""www.microsoft.com/data"">www.microsoft.com/data</a><br/>"
    Response.End
  end if
  'Check content length if needed
  Length = CLng(Request.ServerVariables("Content_Length")) 'Get Content-Length header
  If sizeLimit <> "" Then
    sizeLimit = CLng(sizeLimit) * 1024
    If Length > sizeLimit Then
      Response.Write "Upload size " & FormatNumber(Length, 0) & "B exceeds limit of " & FormatNumber(sizeLimit, 0) & "B<br/>"
      Response.Write "Please correct and <a href=""javascript:history.back(1)"">try again</a>"      
      Response.End
    End If
  End If
End Sub


'Check if version is uptodate
Sub CheckPureUploadVersion(pau_version)
  Dim foundPureUploadVersion
  foundPureUploadVersion = getPureUploadVersion()
  if err or pau_version > foundPureUploadVersion then
    Response.Write "<strong>You don't have latest version of ScriptLibrary/incPureUpload.asp uploaded on the server.</strong><br/>"
    Response.Write "This library is required for the current page. It is fully backwards compatible so old pages will work as well.<br/>"
    Response.End    
  end if
End Sub


'Get fieldname
function pau_Name(FormInfo)
  Dim PosBeg, PosLen
  PosBeg = InStr(FormInfo, "name=")+6
  PosLen = InStr(PosBeg, FormInfo, Chr(34))-PosBeg
  pau_Name = Mid(FormInfo, PosBeg, PosLen)
end function


'Get filename
function pau_FileName(FormInfo)
  Dim PosBeg, PosLen
  PosBeg = InStr(FormInfo, "filename=")+10
  PosLen = InStr(PosBeg, FormInfo, Chr(34))-PosBeg
  pau_FileName = Mid(FormInfo, PosBeg, PosLen)
end function


'Get contentType
function pau_ContentType(FormInfo)
  Dim PosBeg
  PosBeg = InStr(FormInfo, "Content-Type: ")+14
  pau_ContentType = Mid(FormInfo, PosBeg)
end function


'Compatibility with older versions
Sub BuildUploadRequest(RequestBin,UploadDirectory,storeType,sizeLimit,nameConflict)
  pau_BuildUploadRequest RequestBin,UploadDirectory,storeType,sizeLimit,nameConflict,""
End Sub


Sub pau_BuildUploadRequest(RequestBin,UploadDirectory,storeType,sizeLimit,nameConflict,Extensions)
  Dim Boundary, FormInfo, TypeArr, BoundaryArr, BoundaryPos, boundaryEnd, PosBeg, PosEnd, Pos, PosLen, Extension, ExtArr, i
  Dim PosFile, Name, PosBound, FileName, ContentType, Value, ValueBeg, ValueEnd, ValueLen, ExtChk
  'Check content type
  TypeArr = Split(Request.ServerVariables("Content_Type"), ";")
  if Trim(TypeArr(0)) <> "multipart/form-data" then
    Response.Write "<strong>Form was submitted with no ENCTYPE=""multipart/form-data""</strong><br/>"
    Response.Write "Please correct and <a href=""javascript:history.back(1)"">try again</a>"      
    Response.End
  end if
  'Get the boundary
    PosBeg = 1
    PosEnd = InstrB(PosBeg,RequestBin,pau_getByteString(chr(13)))
    if PosEnd = 0 then
    Response.Write "<strong>Form was submitted with no ENCTYPE=""multipart/form-data""</strong><br/>"
    Response.Write "Please correct and <a href=""javascript:history.back(1)"">try again</a>"    
    Response.End
    end if
    boundary = MidB(RequestBin,PosBeg,PosEnd-PosBeg)
    boundaryPos = InstrB(1,RequestBin,boundary)
	boundaryEnd = LenB(RequestBin) - (LenB(boundary) * 2)
    'Get all data inside the boundaries
    Do until (boundaryPos > boundaryEnd)
    'Members variable of objects are put in a dictionary object
    Dim UploadControl
    Set UploadControl = CreateObject("Scripting.Dictionary")
    'Get an object name
    Pos = InstrB(BoundaryPos,RequestBin,pau_getByteString("Content-Disposition"))
    Pos = InstrB(Pos,RequestBin,pau_getByteString("name="))
    PosBeg = Pos+6
    PosEnd = InstrB(PosBeg,RequestBin,pau_getByteString(chr(34)))
    Name = LCase(pau_getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg)))
    PosFile = InstrB(BoundaryPos,RequestBin,pau_getByteString("filename="))
    PosBound = InstrB(PosEnd,RequestBin,boundary)
    'Test if object is of file type
    If  PosFile<>0 AND (PosFile<PosBound) Then
      'Get Filename, content-type and content of file
      PosBeg = PosFile + 10
      PosEnd =  InstrB(PosBeg,RequestBin,pau_getByteString(chr(34)))
      FileName = pau_getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
      FileName = pau_RemoveInvalidChars(Mid(FileName,InStrRev(FileName,"\")+1))
      'Check extension
      Extension = Mid(FileName,InStrRev(FileName,".")+1)
      If Extensions <> "" And FileName <> "" Then
        ExtChk = true
        ExtArr = Split(Extensions, ",")
        For i = 0 to UBound(ExtArr)
          If LCase(Trim(ExtArr(i))) = LCase(Extension) Then
            ExtChk = false
          End If
        Next
        If ExtChk Then
          Response.Write "Filename: " & FileName & "<br/>"
          Response.Write "Filetype is not allowed, only " & Extensions & " are allowed<br/>"
          Response.Write "Please correct and <a href=""javascript:history.back(1)"">try again</a>"
          Response.End
        End If
      End If
      'Add filename to dictionary object
      UploadControl.Add "FileName", FileName
      Pos = InstrB(PosEnd,RequestBin,pau_getByteString("Content-Type:"))
      PosBeg = Pos+14
      PosEnd = InstrB(PosBeg,RequestBin,pau_getByteString(chr(13)))
      'Add content-type to dictionary object
      ContentType = pau_getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
      UploadControl.Add "ContentType",ContentType
      'Get content of object
      PosBeg = PosEnd+4
      PosEnd = InstrB(PosBeg,RequestBin,boundary)-2
      Value = FileName
      ValueBeg = PosBeg-1
      ValueLen = PosEnd-Posbeg
    Else
      'Get content of object
      Pos = InstrB(Pos,RequestBin,pau_getByteString(chr(13)))
      PosBeg = Pos+4
      PosEnd = InstrB(PosBeg,RequestBin,boundary)-2
      Value = pau_getString(MidB(RequestBin,PosBeg,PosEnd-PosBeg))
      ValueBeg = 0
      ValueEnd = 0
    End If
    'Add content to dictionary object
    UploadControl.Add "Value" , Value 
    UploadControl.Add "ValueBeg" , ValueBeg
    UploadControl.Add "ValueLen" , ValueLen 
    'Add dictionary object to main dictionary
    if UploadRequest.Exists(name) then
      UploadRequest(name).Item("Value") = UploadRequest(name).Item("Value") & "," & Value
    else
      UploadRequest.Add name, UploadControl 
    end if    
    'Loop to next object
    BoundaryPos=InstrB(BoundaryPos+LenB(boundary),RequestBin,boundary)
    Loop
  Dim GP_keys, GP_i, GP_curKey, GP_value, GP_valueBeg, GP_valueLen, GP_curPath, GP_FullPath
  Dim GP_CurFileName, GP_FullFileName, fso, GP_BegFolder, GP_RelFolder, GP_FileExist, Begin_Name_Num
  Dim orgUploadDirectory
  if InStr(UploadDirectory,"""") > 0 then 
    on error resume next
    orgUploadDirectory = UploadDirectory
    UploadDirectory = eval(UploadDirectory)  
    if err then
      Response.Write "<strong>Upload folder is invalid</strong><br/><br/>"      
      Response.Write "Upload Folder: " & Trim(orgUploadDirectory) & "<br/>"
      Response.Write "Please correct and <a href=""javascript:history.back(1)"">try again</a>"
      err.clear
      response.End
    end if    
    on error goto 0
  end if  
  GP_keys = UploadRequest.Keys
  for GP_i = 0 to UploadRequest.Count - 1
    GP_curKey = GP_keys(GP_i)
    'Save all uploaded files
    if UploadRequest.Item(GP_curKey).Item("FileName") <> "" then
      GP_value = UploadRequest.Item(GP_curKey).Item("Value")
      GP_valueBeg = UploadRequest.Item(GP_curKey).Item("ValueBeg")
      GP_valueLen = UploadRequest.Item(GP_curKey).Item("ValueLen")
      'Get the path
      if InStr(UploadDirectory,"\") > 0 then
        GP_curPath = UploadDirectory
        if Mid(GP_curPath,Len(GP_curPath),1) <> "\" then
          GP_curPath = GP_curPath & "\"
        end if         
        GP_FullPath = GP_curPath
      else
        if Left(UploadDirectory,1) = "/" then
          GP_curPath = UploadDirectory
        else
          GP_curPath = Request.ServerVariables("PATH_INFO")
          GP_curPath = Trim(Mid(GP_curPath,1,InStrRev(GP_curPath,"/")) & UploadDirectory)
          while InStr(GP_curPath, "/./") > 0
            pos = InStr(GP_curPath, "/./")
            GP_curPath = Trim(Mid(GP_curPath,1,pos) & Mid(GP_curPath,pos+3))
          wend
          while InStr(GP_curPath, "/../") > 0
            pos = InStr(GP_curPath, "/../")
            if pos > 1 then
              GP_curPath = Trim(Mid(GP_curPath,1,InStrRev(GP_curPath, "/", pos-1)) & Mid(GP_curPath,pos+4))
            else
              GP_curPath = Trim(Mid(GP_curPath,1,pos) & Mid(GP_curPath,pos+4))
            end if
          wend
          if Mid(GP_curPath,Len(GP_curPath),1)  <> "/" then
            GP_curPath = GP_curPath & "/"
          end if 
        end if
        GP_FullPath = Trim(Server.mappath(GP_curPath))        
      end if
      if GP_valueLen = 0 then
        Response.Write "<strong>An error has occurred while saving the uploaded file!</strong><br/><br/>"
        Response.Write "Filename: " & Trim(GP_curPath) & UploadRequest.Item(GP_curKey).Item("FileName") & "<br/>"
        Response.Write "The file does not exists or is empty.<br/>"
        Response.Write "Please correct and <a href=""javascript:history.back(1)"">try again</a>"
        response.End
      end if
      'Create a Stream instance
      Dim GP_strm1, GP_strm2
      Set GP_strm1 = Server.CreateObject("ADODB.Stream")
      Set GP_strm2 = Server.CreateObject("ADODB.Stream")
      'Open the stream
      GP_strm1.Open
      GP_strm1.Type = 1 'Binary
      GP_strm2.Open
      GP_strm2.Type = 1 'Binary
      GP_strm1.Write RequestBin
      GP_strm1.Position = GP_ValueBeg
      GP_strm1.CopyTo GP_strm2,GP_ValueLen
      'Create and Write to a File
      GP_CurFileName = UploadRequest.Item(GP_curKey).Item("FileName")      
      GP_FullFileName = GP_FullPath & "\" & GP_CurFileName
      Set fso = CreateObject("Scripting.FileSystemObject")
      pau_AutoCreatePath GP_FullPath
      'Check if the file already exist
      GP_FileExist = false
      If fso.FileExists(GP_FullFileName) Then
        GP_FileExist = true
      End If      
      if nameConflict = "error" and GP_FileExist then
        Response.Write "<strong>The file already exists on the server!</strong><br/><br/>"
        Response.Write "Please correct and <a href=""javascript:history.back(1)"">try again</a>"
        GP_strm1.Close
        GP_strm2.Close
        response.End
      end if
      if ((nameConflict = "over" or nameConflict = "uniq") and GP_FileExist) or (NOT GP_FileExist) then
        if nameConflict = "uniq" and GP_FileExist then
          Begin_Name_Num = 0
          while GP_FileExist    
            Begin_Name_Num = Begin_Name_Num + 1
            GP_FullFileName = Trim(GP_FullPath)& "\" & fso.GetBaseName(GP_CurFileName) & "_" & Begin_Name_Num & "." & fso.GetExtensionName(GP_CurFileName)
            GP_FileExist = fso.FileExists(GP_FullFileName)
          wend  
          UploadRequest.Item(GP_curKey).Item("FileName") = fso.GetBaseName(GP_CurFileName) & "_" & Begin_Name_Num & "." & fso.GetExtensionName(GP_CurFileName)
          UploadRequest.Item(GP_curKey).Item("Value") = UploadRequest.Item(GP_curKey).Item("FileName")
        end if
        on error resume next
        GP_strm2.SaveToFile GP_FullFileName,2
        if err then
          err.clear
          Dim txt_stream, file_bin
          Set txt_stream = fso.CreateTextFile(GP_FullFileName, True)
          file_bin = pau_getString(MidB(RequestBin, GP_ValueBeg+1, GP_ValueLen))
          txt_stream.Write file_bin
          txt_stream.Close
          if err then
            GP_strm1.Close
            GP_strm2.Close
            Response.Write "<strong>An error has occurred while saving uploaded file!</strong><br/><br/>"
            Response.Write "Filename: " & GP_FullFileName & "<br/><br/>"
            if fso.FileExists(GP_FullFileName) then
              Dim f
              Response.Write "The file already exists on the server!<br/>"
              Set f = fso.GetFile(GP_FullFileName)
              Response.Write "Attributes(" & f.attributes & "|" & f.parentfolder.attributes & "): "
              if f.attributes and 1 then
                Response.Write "ReadOnly "
              end if
              if f.attributes and 2 then
                Response.Write "Hidden "
              end if
              if f.attributes and 4 then
                Response.Write "System "
              end if
              if f.attributes and 16 then
                Response.Write "Directory "
              end if
              Response.Write "<br/><br/>"
            end if
            response.End
          end if
        end if
        GP_strm1.Close
        GP_strm2.Close
        if storeType = "path" then
          UploadRequest.Item(GP_curKey).Item("Value") = GP_curPath & UploadRequest.Item(GP_curKey).Item("Value")
        end if
        on error goto 0
      end if
    end if
  next
End Sub


'Create folders if they do not exist
Sub pau_AutoCreatePath(PAU_FullPath)
  Dim FL_fso, FL_EndPos, PAU_NewPath
  Set FL_fso = CreateObject("Scripting.FileSystemObject")  
  if not FL_fso.FolderExists(PAU_FullPath) then
  FL_EndPos = InStrRev(PAU_FullPath,"\")
  if FL_EndPos > 0 then
    PAU_NewPath = Left(PAU_FullPath,FL_EndPos-1)
    pau_AutoCreatePath PAU_NewPath
    on error resume next
    FL_fso.CreateFolder PAU_FullPath
    if err.number <> 0 then
      Response.Write "<strong>Can not create upload folder path: " & PAU_FullPath & "!</strong><br/>"
      Response.Write "Maybe you don't have the proper permissions<br/><br/>"        
      Response.Write "Error # " & CStr(Err.Number) & " " & Err.Description & "<br/><br/>"
      Response.Write "Please correct and <a href=""javascript:history.back(1)"">try again</a>"
      Response.End        
    end if
    on error goto 0
    end if  
  end if
  Set FL_fso = nothing
End Sub


'String to byte string conversion
Function pau_getByteString(StringStr)
  Dim i, char
  For i = 1 to Len(StringStr)
    char = Mid(StringStr,i,1)
    pau_getByteString = pau_getByteString & chrB(AscB(char))
  Next
End Function


'Byte string to string conversion (with double-byte support now)
Function pau_getString(StringBin)
  Dim intCount,get1Byte
  pau_getString = ""
  For intCount = 1 to LenB(StringBin)
    get1Byte = MidB(StringBin,intCount,1)
    pau_getString = pau_getString & chr(AscB(get1Byte)) 
  Next
End Function


'Replacement for the requests
Function UploadFormRequest(name)
  Dim keyName
  keyName = LCase(name)
  if IsObject(UploadRequest) then
    if UploadRequest.Exists(keyName) then
      if UploadRequest.Item(keyName).Exists("Value") then
        UploadFormRequest = UploadRequest.Item(keyName).Item("Value")
      end if  
    end if  
  end if  
End Function


'Invalid characters
'Dollar sign ($) 
'At sign (@) 
'Angle brackets (< >), brackets ([ ]), braces ({ }), and parentheses (( )) 
'Colon (:) and semicolon (;) 
'Equal sign (=) 
'Caret sign (^) 
'Pipe (vertical bar) (|) 
'Asterisk (*) 
'Exclamation point (!) 
'Forward (/) and backward slash (\) 
'Percent sign (%) 
'Question mark (?) 
'Comma (,) 
'Quotation mark (single or double) (' ") 
'Tab 
Function pau_RemoveInvalidChars(str)
  Dim newStr, ci, curChar, Invalid
  Invalid = "$@<>[]{}():;=^|*!/\%?,'""  "
  for ci = 1 to Len(str)
    curChar = Mid(str,ci,1)
    if InStr(Invalid, curChar) = 0 then
      newStr = newStr & curChar
    end if
  next
  pau_RemoveInvalidChars = Trim(newStr)
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
    if UploadRequest.Exists(GP_FieldName) then
      GP_CurFileName = UploadRequest.Item(GP_FieldName).Item("FileName")
      GP_CurContentType = UploadRequest.Item(GP_FieldName).Item("ContentType")
    else  
      GP_CurFileName = ""
      GP_CurContentType = ""
    end if  
    if (GP_CurFileName = "" and GP_CurContentType = "") or (GP_CurFileName <> "" and GP_CurContentType <> "") then
      GP_fieldsStr = GP_fieldsStr & GP_FieldName & "|" & GP_FieldValue & "|"
    end if 
  Next
  if GP_fieldsStr <> "" then
    GP_fieldsStr = Mid(GP_fieldsStr,1,Len(GP_fieldsStr)-1)
  else  
    Response.Write "<strong>An error has occured during record update!</strong><br/><br/>"
    Response.Write "There are no fields to update ...<br/>"
    Response.Write "If the file upload field is the only field on your form, you should make it required.<br/>"
    Response.Write "Please correct and <a href=""javascript:history.back(1)"">try again</a>"
    Response.End
  end if
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
    if UploadRequest.Exists(GP_FieldName) then
      GP_CurFileName = UploadRequest.Item(GP_FieldName).Item("FileName")
      GP_CurContentType = UploadRequest.Item(GP_FieldName).Item("ContentType")    
    else  
      GP_CurFileName = ""
      GP_CurContentType = ""
    end if  
    if (GP_CurFileName = "" and GP_CurContentType = "") or (GP_CurFileName <> "" and GP_CurContentType <> "") then
      GP_columnsStr = GP_columnsStr & GP_ColumnName & "|" & GP_ColumnValue & "|"
    end if 
  Next
  if GP_columnsStr <> "" then
    GP_columnsStr = Mid(GP_columnsStr,1,Len(GP_columnsStr)-1)    
  end if
  FixColumnsForUpload = GP_columnsStr
End Function

</SCRIPT>