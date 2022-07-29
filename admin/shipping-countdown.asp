<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
'==== PAGE HAS BEEN BOOSTRAPPED =======
bootstrapped = "yes"

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
' *** Update Record: set variables

If (CStr(Request("MM_update")) = "FRM_EditTimer" And CStr(Request("MM_recordId")) <> "") Then

  MM_editConnection = MM_bodyartforms_sql_STRING
  MM_editTable = "dbo.TBL_Shipping_Countdown"
  MM_editColumn = "id"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "shipping-countdown.asp"
  MM_fieldsStr  = "timer_name|value|monday|value|tuesday|value|wednesday|value|thursday|value|friday|value|saturday|value|sunday|value|start_time|value|end_time|value|text_message|value"
  MM_columnsStr = "timer_name|',none,''|monday|none,none,NULL|tuesday|none,none,NULL|wednesday|none,none,NULL|thursday|none,none,NULL|friday|none,none,NULL|saturday|none,none,NULL|sunday|none,none,NULL|start_time|',none,''|end_time|',none,''|text_message|',none,''"

  ' create the MM_fields and MM_columns arrays
  MM_fields = Split(MM_fieldsStr, "|")
  MM_columns = Split(MM_columnsStr, "|")
  
  ' set the form values
  For MM_i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_fields(MM_i+1) = CStr(Request.Form(MM_fields(MM_i)))
  Next

  ' append the query string to the redirect URL
  If (MM_editRedirectUrl <> "" And Request.QueryString <> "") Then
    If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
      MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
    Else
      MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
    End If
  End If

End If
%>
<%
' *** Update Record: construct a sql update statement and execute it

If (CStr(Request("MM_update")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  ' create the sql update statement
  MM_editQuery = "update " & MM_editTable & " set "
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
      MM_editQuery = MM_editQuery & ","
    End If
    MM_editQuery = MM_editQuery & MM_columns(MM_i) & " = " & MM_formVal
  Next
  MM_editQuery = MM_editQuery & " where " & MM_editColumn & " = " & MM_recordId

  If (Not MM_abortEdit) Then
    ' execute the update
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
<%
' *** Insert Record: set variables

If (CStr(Request("MM_insert")) = "FRM_AddTimer") Then

  MM_editConnection = MM_bodyartforms_sql_STRING
  MM_editTable = "dbo.TBL_Shipping_Countdown"
  MM_editRedirectUrl = "shipping-countdown.asp"
  MM_fieldsStr  = "timer_name|value|monday|value|tuesday|value|wednesday|value|thursday|value|friday|value|saturday|value|sunday|value|start_time|value|end_time|value|text_message|value"
  MM_columnsStr = "timer_name|',none,''|monday|none,none,NULL|tuesday|none,none,NULL|wednesday|none,none,NULL|thursday|none,none,NULL|friday|none,none,NULL|saturday|none,none,NULL|sunday|none,none,NULL|start_time|',none,''|end_time|',none,''|text_message|',none,''"

  ' create the MM_fields and MM_columns arrays
  MM_fields = Split(MM_fieldsStr, "|")
  MM_columns = Split(MM_columnsStr, "|")
  
  ' set the form values
  For MM_i = LBound(MM_fields) To UBound(MM_fields) Step 2
    MM_fields(MM_i+1) = CStr(Request.Form(MM_fields(MM_i)))
  Next

  ' append the query string to the redirect URL
  If (MM_editRedirectUrl <> "" And Request.QueryString <> "") Then
    If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
      MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
    Else
      MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
    End If
  End If

End If
%>
<%
' *** Delete Record: declare variables

if (CStr(Request("MM_delete")) = "FRM_DeleteTimer" And CStr(Request("MM_recordId")) <> "") Then

  MM_editConnection = MM_bodyartforms_sql_STRING
  MM_editTable = "dbo.TBL_Shipping_Countdown"
  MM_editColumn = "id"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "shipping-countdown.asp"

  ' append the query string to the redirect URL
  If (MM_editRedirectUrl <> "" And Request.QueryString <> "") Then
    If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
      MM_editRedirectUrl = MM_editRedirectUrl
    Else
      MM_editRedirectUrl = MM_editRedirectUrl
    End If
  End If
  
End If
%>
<%
' *** Delete Record: construct a sql delete statement and execute it

If (CStr(Request("MM_delete")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  ' create the sql delete statement
  MM_editQuery = "delete from " & MM_editTable & " where " & MM_editColumn & " = " & MM_recordId

  If (Not MM_abortEdit) Then
    ' execute the delete
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
<%
' *** Insert Record: construct a sql insert statement and execute it

Dim MM_tableValues
Dim MM_dbValues

If (CStr(Request("MM_insert")) <> "") Then

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
Response.Write MM_editQuery
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
<%
Dim rsGetTimers__MMColParam
rsGetTimers__MMColParam = "A"
If (Request("MM_EmptyValue") <> "") Then 
  rsGetTimers__MMColParam = Request("MM_EmptyValue")
End If
%>
<%
Dim rsGetTimers
Dim rsGetTimers_numRows

Set rsGetTimers = Server.CreateObject("ADODB.Recordset")
rsGetTimers.ActiveConnection = MM_bodyartforms_sql_STRING
rsGetTimers.Source = "SELECT * FROM dbo.TBL_Shipping_Countdown ORDER BY id ASC"
rsGetTimers.CursorLocation = 3 'adUseClient
rsGetTimers.LockType = 1 'Read-only records
rsGetTimers.Open()

rsGetTimers_numRows = 0
%>
<%
Dim rsGetONE__MMColParam
rsGetONE__MMColParam = "1"
If (Request.QueryString("ID") <> "") Then 
  rsGetONE__MMColParam = Request.QueryString("ID")
End If
%>
<%
Dim rsGetONE
Dim rsGetONE_numRows

Set rsGetONE = Server.CreateObject("ADODB.Recordset")
rsGetONE.ActiveConnection = MM_bodyartforms_sql_STRING
rsGetONE.Source = "SELECT * FROM dbo.TBL_Shipping_Countdown WHERE id = " + Replace(rsGetONE__MMColParam, "'", "''") + ""
rsGetONE.CursorLocation = 3 'adUseClient
rsGetONE.LockType = 1 'Read-only records
rsGetONE.Open()

rsGetONE_numRows = 0
%>

<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
rsGetTimers_numRows = rsGetTimers_numRows + Repeat1__numRows
%>
<html>
<head>
<title>Manage shipping countdown timer</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body>
<!--#include file="admin_header.asp"-->
<div class="p-3">
<% if request.querystring("Add") = "yes" then %>
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="FRM_AddTimer" id="FRM_AddTimer">
&nbsp;<br>
<table width="50%" border="0" cellspacing="0" cellpadding="1">
    <tr align="left"> 
      <td align="right" valign="middle">Timer name&nbsp;&nbsp;</td>
      <td valign="middle"> <input name="timer_name" type="text" class="adminfields" id="timer_name" size="30"></td>
    </tr>	
    <tr align="left"> 
      <td align="right" valign="middle">Text&nbsp;&nbsp;</td>
      <td valign="middle"> <input name="text_message" type="text" class="adminfields" id="text_message" size="70"></td>
    </tr>		
    <tr align="left"> 
      <td align="right" valign="middle">Start time&nbsp;&nbsp;</td>
      <td valign="middle"> <input name="start_time" type="text" class="adminfields" id="start_time" size="6"></td>
    </tr>
    <tr align="left"> 
      <td align="right" valign="middle">End time&nbsp;&nbsp;</td>
      <td valign="middle"> <input name="end_time" type="text" class="adminfields" id="end_time" size="6"></td>
    </tr>
    <tr align="left"> 
	  <td align="right" valign="texttop">&nbsp;&nbsp;</td>
      <td valign="middle">
	  		<div class="custom-control custom-checkbox d-inline-block">
					<input name="monday" id="monday" class="custom-control-input" type="checkbox" value="1">
					<label class="custom-control-label" for="monday">Monday</label>
			</div>
	  		<div class="custom-control custom-checkbox d-inline-block">
					<input name="tuesday" id="tuesday" class="custom-control-input" type="checkbox" value="1">
					<label class="custom-control-label" for="tuesday">Tuesday</label>
			</div>	
	  		<div class="custom-control custom-checkbox d-inline-block">
					<input name="wednesday" id="wednesday" class="custom-control-input" type="checkbox" value="1">
					<label class="custom-control-label" for="wednesday">Wednesday</label>
			</div>	
	  		<div class="custom-control custom-checkbox d-inline-block">
					<input name="thursday" id="thursday" class="custom-control-input" type="checkbox" value="1">
					<label class="custom-control-label" for="thursday">Thursday</label>
			</div>
	  		<div class="custom-control custom-checkbox d-inline-block">
					<input name="friday" id="friday" class="custom-control-input" type="checkbox" value="1">
					<label class="custom-control-label" for="friday">Friday</label>
			</div>
	  		<div class="custom-control custom-checkbox d-inline-block">
					<input name="saturday" id="saturday" class="custom-control-input" type="checkbox" value="1">
					<label class="custom-control-label" for="saturday">Saturday</label>
			</div>	
	  		<div class="custom-control custom-checkbox d-inline-block">
					<input name="sunday" id="sunday" class="custom-control-input" type="checkbox" value="1">
					<label class="custom-control-label" for="sunday">Sunday</label>
			</div>			
	  </td>
    </tr>
    <tr align="left">
      <td align="right" valign="top">&nbsp;</td>
      <td valign="top"><input class="mt-3" type="submit" name="Submit" value="Submit"></td>
    </tr>
  </table>
<input type="hidden" name="MM_insert" value="FRM_AddTimer">
</form>
<% end if %>


<% if request.querystring("Edit") = "yes" then %>
<%
Dim rsEditTimer__MMColParam
rsEditTimer__MMColParam = "1"
If (Request.QueryString("ID") <> "") Then 
  rsEditTimer__MMColParam = Request.QueryString("ID")
End If
%>
<%
Dim rsEditTimer
Dim rsEditTimer_numRows

Set rsEditTimer = Server.CreateObject("ADODB.Recordset")
rsEditTimer.ActiveConnection = MM_bodyartforms_sql_STRING
rsEditTimer.Source = "SELECT * FROM dbo.TBL_Shipping_Countdown WHERE id = " + Replace(rsEditTimer__MMColParam, "'", "''") + ""
rsEditTimer.CursorLocation = 3 'adUseClient
rsEditTimer.LockType = 1 'Read-only records
rsEditTimer.Open()

rsEditTimer_numRows = 0
%>
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="FRM_EditTimer" id="FRM_EditTimer">
&nbsp;<br>
<table width="50%" border="0" cellspacing="0" cellpadding="1">
    <tr align="left"> 
      <td align="right" valign="middle">Timer name&nbsp;&nbsp;</td>
      <td valign="middle"> <input name="timer_name" type="text" class="adminfields" id="timer_name" size="30" value="<%= rsEditTimer("timer_name")%>"></td>
    </tr>	
    <tr align="left"> 
      <td align="right" valign="middle">Text&nbsp;&nbsp;</td>
      <td valign="middle"> <input name="text_message" type="text" class="adminfields" id="text_message" size="70" value="<%= rsEditTimer("text_message")%>"></td>
    </tr>		
    <tr align="left"> 
      <td align="right" valign="middle">Start time&nbsp;&nbsp;</td>
      <td valign="middle"> <input name="start_time" type="text" class="adminfields" id="start_time" size="6" value="<%= rsEditTimer("start_time")%>"></td>
    </tr>
    <tr align="left"> 
      <td align="right" valign="middle">End time&nbsp;&nbsp;</td>
      <td valign="middle"> <input name="end_time" type="text" class="adminfields" id="end_time" size="6" value="<%= rsEditTimer("end_time")%>"></td>
    </tr>
    <tr align="left"> 
	  <td align="right" valign="texttop">&nbsp;&nbsp;</td>
      <td valign="middle">
	  		<div class="custom-control custom-checkbox d-inline-block">
					<input name="monday" id="monday" class="custom-control-input" type="checkbox" value="1" <% if rsEditTimer("monday") then %>checked<% end if %>>
					<label class="custom-control-label" for="monday">Monday</label>
			</div>
	  		<div class="custom-control custom-checkbox d-inline-block">
					<input name="tuesday" id="tuesday" class="custom-control-input" type="checkbox" value="1" <% if rsEditTimer("tuesday") then %>checked<% end if %>>
					<label class="custom-control-label" for="tuesday">Tuesday</label>
			</div>	
	  		<div class="custom-control custom-checkbox d-inline-block">
					<input name="wednesday" id="wednesday" class="custom-control-input" type="checkbox" value="1" <% if rsEditTimer("wednesday") then %>checked<% end if %>>
					<label class="custom-control-label" for="wednesday">Wednesday</label>
			</div>	
	  		<div class="custom-control custom-checkbox d-inline-block">
					<input name="thursday" id="thursday" class="custom-control-input" type="checkbox" value="1" <% if rsEditTimer("thursday") then %>checked<% end if %>>
					<label class="custom-control-label" for="thursday">Thursday</label>
			</div>
	  		<div class="custom-control custom-checkbox d-inline-block">
					<input name="friday" id="friday" class="custom-control-input" type="checkbox" value="1" <% if rsEditTimer("friday") then %>checked<% end if %>>
					<label class="custom-control-label" for="friday">Friday</label>
			</div>
	  		<div class="custom-control custom-checkbox d-inline-block">
					<input name="saturday" id="saturday" class="custom-control-input" type="checkbox" value="1" <% if rsEditTimer("saturday") then %>checked<% end if %>>
					<label class="custom-control-label" for="saturday">Saturday</label>
			</div>	
	  		<div class="custom-control custom-checkbox d-inline-block">
					<input name="sunday" id="sunday" class="custom-control-input" type="checkbox" value="1" <% if rsEditTimer("sunday") then %>checked<% end if %>>
					<label class="custom-control-label" for="sunday">Sunday</label>
			</div>			
	  </td>
    </tr>
  
  <tr align="left">
    <td align="right" valign="top">&nbsp;</td>
    <td valign="top"><input type="submit" name="Submit3" value="Submit"></td>
  </tr>
</table>

<input type="hidden" name="MM_update" value="FRM_EditTimer">
<input type="hidden" name="MM_recordId" value="<%= request.querystring("ID") %>">
</form>
<% end if %>
<% if request.querystring("Delete") = "yes" then %>
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="FRM_DeleteTimer" id="FRM_DeleteTimer">
  <div class="alert alert-danger">Confirm deletion <input class="btn btn-sm btn-danger ml-4" type="submit" name="Submit2" value="Delete">
  </div>
    <input type="hidden" name="MM_delete" value="FRM_DeleteTimer">
<input type="hidden" name="MM_recordId" value="<%= request.querystring("ID") %>">
</form>
<% end if %>

<h5>Shipping Countdown Timers</h5>
<table class="table table-striped table-borderless table-hover">
<thead class="thead-dark">
  <tr>
	<th>&nbsp;</th>
    <th width="15%">Name
      <a class="ml-3 text-success" href="shipping-countdown.asp?Add=yes"><i class="fa fa-plus-circle mr-1"></i>Add New</a>
	</th>
    <th width="15%">Start Time</th>
    <th width="15%">End Time </th>
    <th width="35%">Text Message</th>
  </tr>
</thead>
  <% 
While ((Repeat1__numRows <> 0) AND (NOT rsGetTimers.EOF)) 
%>
    <tr>
      <td>
		  <a class="btn btn-sm btn-danger mr-2" href="shipping-countdown.asp?Delete=yes&ID=<%=(rsGetTimers.Fields.Item("id").Value)%>"><i class="fa fa-trash-alt"></i></a>
		  <a href="shipping-countdown.asp?Edit=yes&ID=<%=(rsGetTimers.Fields.Item("id").Value)%>" target="_top" class="LeftNavLinks"><%=(rsGetTimers.Fields.Item("timer_name").Value)%></a>
	  </td>
      <td><%=(rsGetTimers.Fields.Item("timer_name").Value)%></td>
	  <td><%=(rsGetTimers.Fields.Item("start_time").Value)%></td>
	  <td><%=(rsGetTimers.Fields.Item("end_time").Value)%></td>
	  <td><%=(rsGetTimers.Fields.Item("text_message").Value)%></td>
    </tr>
    <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsGetTimers.MoveNext()
Wend
%>
</table>
</div>
</body>
</html>

<%
rsGetTimers.Close()
Set rsGetTimers = Nothing
%>
<%
rsGetONE.Close()
Set rsGetONE = Nothing
%>

