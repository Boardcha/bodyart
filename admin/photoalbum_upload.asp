<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../ScriptLibrary/incPureUpload.asp" -->
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
'*** Pure ASP File Upload 2.2.1
Dim GP_uploadAction,UploadQueryString
PureUploadSetup
If (CStr(Request.QueryString("GP_upload")) <> "") Then
  Dim pau_thePath,pau_Extensions,pau_Form,pau_Redirect,pau_storeType,pau_sizeLimit,pau_nameConflict,pau_requireUpload,pau_minWidth,pau_minHeight,pau_maxWidth,pau_maxHeight,pau_saveWidth,pau_saveHeight,pau_timeout,pau_progressBar,pau_progressWidth,pau_progressHeight
  pau_thePath = """../gallery/uploads"""
  pau_Extensions = "GIF,JPG,JPEG"
  pau_Form = "UploadPhoto"
  pau_Redirect = ""
  pau_storeType = "file"
  pau_sizeLimit = ""
  pau_nameConflict = "uniq"
  pau_requireUpload = "false"
  pau_minWidth = ""
  pau_minHeight = "" 
  pau_maxWidth = ""
  pau_maxHeight = ""
  pau_saveWidth = ""
  pau_saveHeight = ""
  pau_timeout = "600"
  pau_progressBar = ""
  pau_progressWidth = "300"
  pau_progressHeight = "100"
  
  Dim RequestBin, UploadRequest
  CheckPureUploadVersion 2.21
  ProcessUpload pau_thePath,pau_Extensions,pau_Redirect,pau_storeType,pau_sizeLimit,pau_nameConflict,pau_requireUpload,pau_minWidth,pau_minHeight,pau_maxWidth,pau_maxHeight,pau_saveWidth,pau_saveHeight,pau_timeout
end if
%>
<%
' *** Edit Operations: (Modified for File Upload) declare variables

Dim MM_editAction
Dim MM_abortEdit
Dim MM_editQuery
Dim MM_editCmd

Dim MM_editConnection
Dim MM_editTable
Dim MM_editRedirectUrl
Dim MM_editColumn
Dim MM_recordId

Dim MM_fieldsStr
Dim MM_columnsStr
Dim MM_fields
Dim MM_columns
Dim MM_typeArray
Dim MM_formVal
Dim MM_delim
Dim MM_altVal
Dim MM_emptyVal
Dim MM_i

MM_editAction = CStr(Request.ServerVariables("SCRIPT_NAME"))
If (UploadQueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Server.HTMLEncode(UploadQueryString)
End If

' boolean to abort record edit
MM_abortEdit = false

' query string to execute
MM_editQuery = ""
%>
<%
' *** Insert Record: (Modified for File Upload) set variables

If (CStr(UploadFormRequest("MM_insert")) = "UploadPhoto") Then

  MM_editConnection = MM_bodyartforms_sql_STRING
  MM_editTable = "dbo.TBL_PhotoGallery"
  MM_editRedirectUrl = "photoalbum_upload.asp"
  MM_fieldsStr  = "FILE1|value|type|value|description|value"
  MM_columnsStr = "filename|',none,''|type|',none,''|description|',none,''"

  ' create the MM_fields and MM_columns arrays
  MM_fields = Split(MM_fieldsStr, "|")
  MM_columns = Split(MM_columnsStr, "|")
  
  ' set the form values
  For MM_i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_fields(MM_i+1) = CStr(UploadFormRequest(MM_fields(MM_i)))
  Next

  ' append the query string to the redirect URL
  If (MM_editRedirectUrl <> "" And UploadQueryString <> "") Then
    If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0 And UploadQueryString <> "") Then
      MM_editRedirectUrl = MM_editRedirectUrl & "?" & UploadQueryString
    Else
      MM_editRedirectUrl = MM_editRedirectUrl & "&" & UploadQueryString
    End If
  End If

End If
%>
<%
' *** Insert Record: (Modified for File Upload) construct a sql insert statement and execute it

Dim MM_tableValues
Dim MM_dbValues

If (CStr(UploadFormRequest("MM_insert")) <> "") Then

  ' create the sql insert statement
  MM_tableValues = ""
  MM_dbValues = ""
  For MM_i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_formVal = MM_fields(MM_i+1)
    MM_typeArray = Split(MM_columns(MM_i+1),",")
    MM_delim = MM_typeArray(0)
    If (MM_delim = "none") Then MM_delim = ""
    MM_altVal = MM_typeArray(1)
    If (MM_altVal = "none") Then MM_altVal = ""
    MM_emptyVal = MM_typeArray(2)
    If (MM_emptyVal = "none") Then MM_emptyVal = ""
    If (MM_formVal = "") Then
      MM_formVal = MM_emptyVal
    Else
      If (MM_altVal <> "") Then
        MM_formVal = MM_altVal
      ElseIf (MM_delim = "'") Then  ' escape quotes
        MM_formVal = "'" & Replace(MM_formVal,"'","''") & "'"
      Else
        MM_formVal = MM_delim + MM_formVal + MM_delim
      End If
    End If
    If (MM_i <> LBound(MM_fields)) Then
      MM_tableValues = MM_tableValues & ","
      MM_dbValues = MM_dbValues & ","
    End If
    MM_tableValues = MM_tableValues & MM_columns(MM_i)
    MM_dbValues = MM_dbValues & MM_formVal
  Next
  MM_editQuery = "insert into " & MM_editTable & " (" & MM_tableValues & ") values (" & MM_dbValues & ")"

  If (Not MM_abortEdit) Then
    ' execute the insert
    Set MM_editCmd = Server.CreateObject("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_editConnection
    MM_editCmd.CommandText = MM_editQuery
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    If (MM_editRedirectUrl <> "") Then
      Response.Redirect(MM_editRedirectUrl)
    End If
  End If

End If
%>
<html>
<head>
<link rel="stylesheet" type="text/css" href="../includes/nav.css" />
<title>Upload photo album pics</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="javascript" src="../ScriptLibrary/incPureUpload.js"></script>
</head>
<body bgcolor="#666699" topmargin="0" text="#CCCCCC" link="#CCCCCC" vlink="#CCCCCC">
<!--#include file="admin_header.asp"-->
<span class="adminheader">Upload photo to gallery</span>
<font face="Verdana" size="2"></font>
<p>
<FORM action="<%=MM_editAction%>" METHOD="POST" ENCTYPE="multipart/form-data" name="UploadPhoto" onSubmit="checkFileUpload(this,'GIF,JPG,JPEG',false,'','','','','','','');return document.MM_returnValue">
      <p><span class="pricegauge">Main photo:</span>
        <INPUT NAME="FILE1" TYPE="FILE" class="adminfields" onChange="checkOneFileUpload(this,'GIF,JPG,JPEG',false,'','','','','','','')" SIZE="40">
        <BR>
          <span class="pricegauge">Thumbnail image:</span>
        <INPUT NAME="FILE2" TYPE="FILE" class="adminfields" onChange="checkOneFileUpload(this,'GIF,JPG,JPEG',false,'','','','','','','')" SIZE="40">
      </p>
      <p><span class="pricegauge">Type:</span> 
        <select name="type" class="adminfields" id="type">
          <option selected value="None">SELECT TYPE</option>
          <option value="Small">Plugs - 0g &amp; smaller</option>
          <option value="Medium">Plugs - 00g thru 5/8&quot;</option>
          <option value="Large">Plugs - larger than 5/8"</option>
          <option value="Labret">Labret</option>
          <option value="Navel">Navel</option>
		  <option value="Nipple">Nipple</option>
          <option value="Septum">Septum</option>
          <option value="Tongue">Tongue</option>
          <option value="Surface">Surface piercing</option>
          <option value="Eyebrow">Eyebrow</option>
          <option value="Multiple facial piercings">Multiple facial piercings</option>
          <option value="Multiple ear piercings">Multiple ear piercings</option>
		  <option value="Misc">Misc</option>
        </select>
</p>
      <p><span class="pricegauge">Description:</span><br>
        <textarea name="description" cols="40" rows="3" class="adminfields" id="description"></textarea>
        <br>
        <BR>
        <INPUT TYPE=SUBMIT class="adminfields" VALUE="Upload!">
  </p>

            <input type="hidden" name="MM_insert" value="UploadPhoto">
</FORM>
<br>
<p>&nbsp;</p>
</body>
</html>