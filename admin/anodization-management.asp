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

If (CStr(Request("MM_update")) = "FRM_Edit" And CStr(Request("MM_recordId")) <> "") Then

  MM_editConnection = MM_bodyartforms_sql_STRING
  MM_editTable = "dbo.TBL_Anodization_Colors_Pricing"
  MM_editColumn = "anodID"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "anodization-management.asp"
  MM_fieldsStr  = "color_name|value|standard_anodization|value|high_voltage|value|base_price|value|multiple_discount_price|value"
  MM_columnsStr = "color_name|',none,''|standard_anodization|',none,''|high_voltage|',none,''|base_price|,none,none,NULL|multiple_discount_price|none,none,NULL"

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
	If MM_formVal ="'on'" Then MM_formVal = 1
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

If (CStr(Request("MM_insert")) = "FRM_Add") Then

  MM_editConnection = MM_bodyartforms_sql_STRING
  MM_editTable = "dbo.TBL_Anodization_Colors_Pricing"
  MM_editRedirectUrl = "anodization-management.asp"
  MM_fieldsStr  = "color_name|value|standard_anodization|value|high_voltage|value|base_price|value|multiple_discount_price|value"
  MM_columnsStr = "color_name|',none,''|standard_anodization|',none,''|high_voltage|',none,''|base_price|,none,none,NULL|multiple_discount_price|none,none,NULL"

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

if (CStr(Request("MM_delete")) = "FRM_Delete" And CStr(Request("MM_recordId")) <> "") Then

  MM_editConnection = MM_bodyartforms_sql_STRING
  MM_editTable = "dbo.TBL_Anodization_Colors_Pricing"
  MM_editColumn = "anodID"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "anodization-management.asp"

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
	If MM_formVal ="'on'" Then MM_formVal = 1
    MM_dbValues = MM_dbValues & MM_formVal
  Next
  MM_editQuery = "insert into " & MM_editTable & " (" & MM_tableValues & ") values (" & MM_dbValues & ")"
'Response.Write MM_editQuery
'Response.End

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
Dim rsGetAnodization__MMColParam
rsGetAnodization__MMColParam = "A"
If (Request("MM_EmptyValue") <> "") Then 
  rsGetAnodization__MMColParam = Request("MM_EmptyValue")
End If
%>
<%
Dim rsGetAnodization
Dim rsGetAnodization_numRows

Set rsGetAnodization = Server.CreateObject("ADODB.Recordset")
rsGetAnodization.ActiveConnection = MM_bodyartforms_sql_STRING
rsGetAnodization.Source = "SELECT * FROM dbo.TBL_Anodization_Colors_Pricing"
rsGetAnodization.CursorLocation = 3 'adUseClient
rsGetAnodization.LockType = 1 'Read-only records
rsGetAnodization.Open()

rsGetAnodization_numRows = 0
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
rsGetONE.Source = "SELECT * FROM dbo.TBL_Anodization_Colors_Pricing WHERE anodID = " + Replace(rsGetONE__MMColParam, "'", "''") + ""
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
rsGetAnodization_numRows = rsGetAnodization_numRows + Repeat1__numRows
%>
<html>
<head>
<title>Anodization color & voltage pricing</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body>
<!--#include file="admin_header.asp"-->
<div class="p-3">
<% if request.querystring("Add") = "yes" then %>
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="FRM_Add" id="FRM_Add">
&nbsp;<br>
<table width="50%" border="0" cellspacing="0" cellpadding="1">
    <tr align="left"> 
      <td align="right" valign="middle" class="pricegauge">Color name &nbsp;</td>
      <td valign="middle" class="pricegauge"> <input name="color_name" type="text" class="adminfields" id="color_name" size="30"></td>
    </tr>
    <tr align="left"> 
      <td align="right" valign="middle" class="pricegauge">Standard anodization&nbsp;&nbsp; </td>
      <td valign="middle"><input name="standard_anodization" type="checkbox" value="1"></td>
    </tr>
	    <tr align="left">
      <td align="right" valign="top" class="pricegauge">High voltage&nbsp;&nbsp;</td>
      <td valign="middle" class="pricegauge"><input name="high_voltage" type="checkbox" value="1" class="adminfields" id="high_voltage"></td>
      </td>
    </tr>
	<tr align="left">
      <td align="right" valign="top" class="pricegauge">Base price&nbsp;&nbsp;</td>
      <td valign="top" class="pricegauge"><input name="base_price" type="text" class="adminfields" id="base_price" value="0" size="6"></td>
    </tr>
    <tr align="left">
      <td align="right" valign="top" class="pricegauge">&nbsp;</td>
      <td valign="top" class="pricegauge"><input type="submit" name="Submit" value="Submit"></td>
    </tr>
  </table>



<input type="hidden" name="MM_insert" value="FRM_Add">
</form>
<% end if %>
<% if request.querystring("Edit") = "yes" then %>
<%
Dim rsEditCompany__MMColParam
rsEditCompany__MMColParam = "1"
If (Request.QueryString("ID") <> "") Then 
  rsEditCompany__MMColParam = Request.QueryString("ID")
End If
%>
<%
Dim rsEditCompany
Dim rsEditCompany_numRows

Set rsEditCompany = Server.CreateObject("ADODB.Recordset")
rsEditCompany.ActiveConnection = MM_bodyartforms_sql_STRING
rsEditCompany.Source = "SELECT * FROM dbo.TBL_Anodization_Colors_Pricing WHERE anodID = " + Replace(rsEditCompany__MMColParam, "'", "''") + ""
rsEditCompany.CursorLocation = 3 'adUseClient
rsEditCompany.LockType = 1 'Read-only records
rsEditCompany.Open()

rsEditCompany_numRows = 0
%>
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="FRM_Edit" id="FRM_Edit">
&nbsp;<br>
<table width="50%" border="0" cellspacing="0" cellpadding="1">
    <tr align="left"> 
      <td align="right" valign="middle" class="pricegauge">Color name &nbsp;</td>
      <td valign="middle" class="pricegauge"> <input name="color_name" type="text" class="adminfields" id="color_name" value="<%=(rsGetONE.Fields.Item("color_name").Value)%>" size="30"></td>
    </tr>
    <tr align="left"> 
      <td align="right" valign="middle" class="pricegauge">Standard anodization&nbsp;&nbsp; </td>
      <td valign="middle" class="pricegauge"><input name="standard_anodization" type="checkbox" <% if rsGetONE.Fields.Item("standard_anodization").Value = True then %>checked<% end if %>></td>
    </tr>
	    <tr align="left">
      <td align="right" valign="top" class="pricegauge">High voltage&nbsp;&nbsp;</td>
      <td valign="middle" class="pricegauge"><input name="high_voltage" type="checkbox" <% if rsGetONE.Fields.Item("high_voltage").Value = True then %>checked<% end if %>></td>
      </td>
    </tr>
	<tr align="left">
      <td align="right" valign="top" class="pricegauge">Base price&nbsp;&nbsp;</td>
      <td valign="top" class="pricegauge"><input name="base_price" type="text" class="adminfields" id="base_price" value="<%=(rsGetONE.Fields.Item("base_price").Value)%>" size="6"></td>
    </tr>
    <tr align="left">
      <td align="right" valign="top" class="pricegauge">&nbsp;</td>
      <td valign="top" class="pricegauge"><input type="submit" name="Submit" value="Submit"></td>
    </tr>
  </table>
  

<input type="hidden" name="MM_update" value="FRM_Edit">
<input type="hidden" name="MM_recordId" value="<%= request.querystring("ID") %>">
</form>
<% end if %>
<% if request.querystring("Delete") = "yes" then %>
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="FRM_Delete" id="FRM_Delete">
  <div class="alert alert-danger">Confirm deletion <input class="btn btn-sm btn-danger ml-4" type="submit" name="Submit2" value="Delete">
  </div>
    <input type="hidden" name="MM_delete" value="FRM_Delete">
<input type="hidden" name="MM_recordId" value="<%= request.querystring("ID") %>">
</form>
<% end if %>
<h5>Anodization color & voltage pricing</h5>
<table class="table table-striped table-borderless table-hover">
<thead class="thead-dark">
  <tr>
    <th width="15%">Color
      <a class="ml-3 text-success" href="anodization-management.asp?Add=yes"><i class="fa fa-plus-circle mr-1"></i>Add New</a></th>
    <th width="5%">Standard</th>
    <th width="30%">High voltage </th>
    <th width="25%">Price</th>
  </tr>
</thead>
  <% 
While ((Repeat1__numRows <> 0) AND (NOT rsGetAnodization.EOF)) 
%>
    <tr>
      <td class="font-weight-bold">
      <a class="btn btn-sm btn-danger mr-2" href="anodization-management.asp?Delete=yes&ID=<%=(rsGetAnodization.Fields.Item("anodID").Value)%>"><i class="fa fa-trash-alt"></i></a>
      <a href="anodization-management.asp?Edit=yes&ID=<%=(rsGetAnodization.Fields.Item("anodID").Value)%>" target="_top" class="LeftNavLinks"><%=(rsGetAnodization.Fields.Item("color_name").Value)%></a></td>
      <td><% If rsGetAnodization.Fields.Item("standard_anodization").Value = true Then%><span class="text-success font-weight-bold">Yes</span><%Else%>No<%End If%></td>
      <td><% If rsGetAnodization.Fields.Item("high_voltage").Value = true Then%><span class="text-success font-weight-bold">Yes</span><%Else%>No<%End If%></td>
      <td><strong><%=(rsGetAnodization.Fields.Item("base_price").Value)%></strong></td>
    </tr>
    <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsGetAnodization.MoveNext()
Wend
%>
</table>
</div>
</body>
</html>

<%
rsGetAnodization.Close()
Set rsGetAnodization = Nothing
%>
<%
rsGetONE.Close()
Set rsGetONE = Nothing
%>

