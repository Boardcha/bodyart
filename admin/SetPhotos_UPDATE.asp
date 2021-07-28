<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
' *** Edit Operations: declare variables

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
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Server.HTMLEncode(Request.QueryString)
End If

' boolean to abort record edit
MM_abortEdit = false

' query string to execute
MM_editQuery = ""
%>
<%
Dim rsSetPhotos__MMColParam
rsSetPhotos__MMColParam = "1"
If (Request("MM_EmptyValue") <> "") Then 
  rsSetPhotos__MMColParam = Request("MM_EmptyValue")
End If
%>
<%
Dim rsSetPhotos
Dim rsSetPhotos_cmd
Dim rsSetPhotos_numRows

Set rsSetPhotos_cmd = Server.CreateObject ("ADODB.Command")
rsSetPhotos_cmd.ActiveConnection = MM_bodyartforms_sql_STRING
rsSetPhotos_cmd.CommandText = "SELECT PhotoID, ProductID, DetailID, filename, [description] FROM dbo.TBL_PhotoGallery WHERE status = ? AND DetailID = 0 ORDER BY PhotoID ASC" 
rsSetPhotos_cmd.Prepared = true
rsSetPhotos_cmd.Parameters.Append rsSetPhotos_cmd.CreateParameter("param1", 5, 1, -1, rsSetPhotos__MMColParam) ' adDouble

Set rsSetPhotos = rsSetPhotos_cmd.Execute
rsSetPhotos_numRows = 0
%>

<%
Dim Repeat5__numRows
Dim Repeat5__index

Repeat5__numRows = 10
Repeat5__index = 0
rsSetPhotos_numRows = rsSetPhotos_numRows + Repeat5__numRows
%>
<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

Dim rsSetPhotos_total
Dim rsSetPhotos_first
Dim rsSetPhotos_last

' set the record count
rsSetPhotos_total = rsSetPhotos.RecordCount

' set the number of rows displayed on this page
If (rsSetPhotos_numRows < 0) Then
  rsSetPhotos_numRows = rsSetPhotos_total
Elseif (rsSetPhotos_numRows = 0) Then
  rsSetPhotos_numRows = 1
End If

' set the first and last displayed record
rsSetPhotos_first = 1
rsSetPhotos_last  = rsSetPhotos_first + rsSetPhotos_numRows - 1

' if we have the correct record count, check the other stats
If (rsSetPhotos_total <> -1) Then
  If (rsSetPhotos_first > rsSetPhotos_total) Then
    rsSetPhotos_first = rsSetPhotos_total
  End If
  If (rsSetPhotos_last > rsSetPhotos_total) Then
    rsSetPhotos_last = rsSetPhotos_total
  End If
  If (rsSetPhotos_numRows > rsSetPhotos_total) Then
    rsSetPhotos_numRows = rsSetPhotos_total
  End If
End If
%>
<%
' *** Recordset Stats: if we don't know the record count, manually count them

If (rsSetPhotos_total = -1) Then

  ' count the total records by iterating through the recordset
  rsSetPhotos_total=0
  While (Not rsSetPhotos.EOF)
    rsSetPhotos_total = rsSetPhotos_total + 1
    rsSetPhotos.MoveNext
  Wend

  ' reset the cursor to the beginning
  If (rsSetPhotos.CursorType > 0) Then
    rsSetPhotos.MoveFirst
  Else
    rsSetPhotos.Requery
  End If

  ' set the number of rows displayed on this page
  If (rsSetPhotos_numRows < 0 Or rsSetPhotos_numRows > rsSetPhotos_total) Then
    rsSetPhotos_numRows = rsSetPhotos_total
  End If

  ' set the first and last displayed record
  rsSetPhotos_first = 1
  rsSetPhotos_last = rsSetPhotos_first + rsSetPhotos_numRows - 1
  
  If (rsSetPhotos_first > rsSetPhotos_total) Then
    rsSetPhotos_first = rsSetPhotos_total
  End If
  If (rsSetPhotos_last > rsSetPhotos_total) Then
    rsSetPhotos_last = rsSetPhotos_total
  End If

End If
%>
<%
' *** FX Update Multiple Records in FRM_UpdatePhotos
If (rsSetPhotos_first <> "") Then upd_rsSetPhotos = rsSetPhotos_first-1 Else upd_rsSetPhotos = 0 End If ' counter
If (cStr(Request.Form("SubmitActive")) <> "") Then
  FX_sqlerror = ""
  FX_updredir = "SetPhotos.asp"
  tmp = "ADODB.Command"
  Set update_Multi = Server.CreateObject(tmp)
  update_Multi.ActiveConnection = MM_bodyartforms_sql_STRING
  For N = upd_rsSetPhotos+1 To rsSetPhotos_total
      If (Request.Form("DetailID"&N) <> "") Then s1 = Replace(Request.Form("DetailID"&N),"'","''") Else s1 = "0" End If
      If (Request.Form("ProductID"&N) <> "") Then s2 = Replace(Request.Form("ProductID"&N),"'","''") Else s2 = "0" End If
      If (Request.Form("fx_updmatch"&N) <> "") Then sw = Replace(Request.Form("fx_updmatch"&N),"'","''") Else sw = "0" End If
    On Error Resume Next
      update_Multi.CommandText = "UPDATE dbo.TBL_PhotoGallery SET DetailID="+s1+", ProductID="+s2+" WHERE PhotoID="+sw+""
      update_Multi.Execute
    If (Err.Description <> "") Then
      FX_sqlerror = FX_sqlerror & "Row " & N & ": " & Err.Description & "<br><br>"
    End If
  Next
  update_Multi.ActiveConnection.Close
  thispath = cStr(Request.ServerVariables("SCRIPT_NAME"))
  If (FX_updredir = "") Then FX_updredir = Mid(thispath, InstrRev(thispath, "/")+1) End If
  If (Request.QueryString <> "") Then
    ch = "&"
    If (InStr(FX_updredir,"?") = 0) Then ch = "?" End If
    FX_updredir = FX_updredir & ch & Request.QueryString
  End If
  If (FX_sqlerror <> "") Then
    Response.Write("<font color=""red"">"&FX_sqlerror&"</font>")
  Else Response.Redirect(FX_updredir) End If
End If
%>
<html>
<head>
<title>Set photo gauges</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body class="mainbkgd">
</body>
</html>
