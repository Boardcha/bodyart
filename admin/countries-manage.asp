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
' *** Delete Record: declare variables

if (CStr(Request("MM_delete")) = "FRM_DeleteCompany" And CStr(Request("MM_recordId")) <> "") Then

  MM_editConnection = MM_bodyartforms_sql_STRING
  MM_editTable = "dbo.TBL_Countries"
  MM_editColumn = "CountryID"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "countries-manage.asp"

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
Dim rsGetCountries
Dim rsGetCountries_numRows

Set rsGetCountries = Server.CreateObject("ADODB.Recordset")
rsGetCountries.ActiveConnection = MM_bodyartforms_sql_STRING
rsGetCountries.Source = "SELECT * FROM dbo.TBL_Countries ORDER BY Country ASC"
rsGetCountries.CursorLocation = 3 'adUseClient
rsGetCountries.LockType = 1 'Read-only records
rsGetCountries.Open()

rsGetCountries_numRows = 0
%>
<%
Dim rsEditCountry__MMColParam
rsEditCountry__MMColParam = "1"
If (Request.QueryString("ID") <> "") Then 
  rsEditCountry__MMColParam = Request.QueryString("ID")
End If
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
rsGetCountries_numRows = rsGetCountries_numRows + Repeat1__numRows
%>
<html>
<head>
<title>Manage countries</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body>
<!--#include file="admin_header.asp"-->
<div class="p-3">

<% if request.querystring("Add") = "yes" then %>
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="FRM_AddCompany" id="FRM_AddCompany">
  <h5>Add a new country</h5>
<table class="table table-sm table-borderless w-25">
    <tr> 
      <td>Country name</td>
      <td> <input name="country" type="text" class="form-control form-control-sm" id="country"></td>
    </tr>
    
    <tr>
      <td>Shipping type&nbsp; </td>
      <td><label>
        <select name="ShippingType" class="form-control form-control-sm" id="ShippingType">
          <option value="0">USA</option>
          <option value="1" selected>International</option>
                </select>
      </label></td>
    </tr>
    <tr>
      <td>2 letter country code</td>
      <td><input name="Country_UPSCode" type="text" class="form-control form-control-sm" id="Country_UPSCode" size="4" maxlength="4"></td>
    </tr>
    <tr>
      <td>Display in country drop down</td>
      <td><input type="radio" name="Display" value="1">
      Yes 
        <input class="ml-3" name="Display" type="radio" value="0" checked>
      No</td>
    </tr>
    <tr>
      <td>Show as origin country in admin products</td>
      <td><input type="radio" name="origin_toggle" value="1">
        Yes
        <input class="ml-3" name="origin_toggle" type="radio" value="0" checked>
        No</td>
        
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td><input class="btn btn-purple" type="submit" name="Submit" value="Submit">
      <input type="hidden" name="dateadded" value="<%= date() %>"></td>
    </tr>
  </table>

<input type="hidden" name="MM_insert" value="FRM_AddCompany">
</form>
<%
' *** Insert Record: set variables

If (CStr(Request("MM_insert")) = "FRM_AddCompany") Then

  MM_editConnection = MM_bodyartforms_sql_STRING
  MM_editTable = "dbo.TBL_Countries"
  MM_editRedirectUrl = "countries-manage.asp"
  MM_fieldsStr  = "country|value|ShippingType|value|Country_UPSCode|value|Display|value|origin_toggle|value"
  MM_columnsStr = "Country|',none,''|CountryType|',none,''|Country_UPSCode|',none,''|Display|',none,''|origin_toggle|',none,''"

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
      MM_editRedirectUrl = MM_editRedirectUrl
    Else
      MM_editRedirectUrl = MM_editRedirectUrl
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
<% end if %>
<% if request.querystring("Edit") = "yes" then %>
<%
Dim rsEditCountry
Dim rsEditCountry_numRows

Set rsEditCountry = Server.CreateObject("ADODB.Recordset")
rsEditCountry.ActiveConnection = MM_bodyartforms_sql_STRING
rsEditCountry.Source = "SELECT * FROM dbo.TBL_Countries WHERE CountryID = " + Replace(rsEditCountry__MMColParam, "'", "''") + ""
rsEditCountry.CursorLocation = 3 'adUseClient
rsEditCountry.LockType = 1 'Read-only records
rsEditCountry.Open()

rsEditCountry_numRows = 0
%>
<h5>Edit country</h5>
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="FRM_EditCompany" id="FRM_EditCompany">
<table class="table table-sm table-borderless w-25">
    <tr> 
      <td>Country name</td>
      <td> <input name="country" type="text" class="form-control form-control-sm" id="country" value="<%=(rsEditCountry.Fields.Item("country").Value)%>" size="30"></td>
    </tr>
    <tr>
      <td>Shipping type&nbsp; </td>
      <td><label>
        <select name="ShippingType" class="form-control form-control-sm" id="ShippingType">
          <option value="0">USA</option>
          <option value="1" selected>International</option>
        </select>
      </label></td>
    </tr>
    <tr>
      <td>2 letter country code</td>
      <td><input name="Country_UPSCode" type="text" class="form-control form-control-sm" id="Country_UPSCode" value="<%=(rsEditCountry.Fields.Item("Country_UPSCode").Value)%>" size="4" maxlength="4"></td>
    </tr>
    <tr>
      <td>Display in country drop down</td>
      <td><input type="radio" name="Display" id="radio3" value="1" <% if rsEditCountry.Fields.Item("Display").Value = 1 Then %>checked<% end if %>>
        Yes
        <input class="ml-3" name="Display" type="radio" id="radio4" value="0" <% if rsEditCountry.Fields.Item("Display").Value = 0 Then %>checked<% end if %>>
        No</td>
        
    </tr>
    <tr>
      <td>Show as origin country in admin products</td>
      <td><input type="radio" name="origin_toggle" value="1" <% if rsEditCountry.Fields.Item("origin_toggle").Value = 1 Then %>checked<% end if %>>
        Yes
        <input class="ml-3" name="origin_toggle" type="radio" value="0" <% if rsEditCountry.Fields.Item("origin_toggle").Value = 0 Then %>checked<% end if %>>
        No</td>
        
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td><input class="btn btn-purple" type="submit" name="Submit" value="Submit"></td>
    </tr>
  </table>

<input type="hidden" name="MM_update" value="FRM_EditCompany">
<input type="hidden" name="MM_recordId" value="<%= rsEditCountry.Fields.Item("CountryID").Value %>">
</form>
<%
' *** Update Record: set variables

If (CStr(Request("MM_update")) = "FRM_EditCompany" And CStr(Request("MM_recordId")) <> "") Then

  MM_editConnection = MM_bodyartforms_sql_STRING
  MM_editTable = "dbo.TBL_Countries"
  MM_editColumn = "CountryID"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "countries-manage.asp"
  MM_fieldsStr  = "country|value|ShippingType|value|Country_UPSCode|value|Display|value|origin_toggle|value"
  MM_columnsStr = "Country|',none,''|CountryType|',none,''|Country_UPSCode|',none,''|Display|',none,''|origin_toggle|',none,''"

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
      MM_editRedirectUrl = MM_editRedirectUrl
    Else
      MM_editRedirectUrl = MM_editRedirectUrl
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
<% end if %>
<% if request.querystring("Delete") = "yes" then %>
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="FRM_DeleteCompany" id="FRM_DeleteCompany">
  <div class="alert alert-danger">
  Confirm deletion<br/> <input class="btn btn-sm btn-danger mt-2" type="submit" name="Submit2" value="DELETE">
</div>

    <input type="hidden" name="MM_delete" value="FRM_DeleteCompany">
    <input type="hidden" name="MM_recordId" value="<%= request.querystring("ID") %>">

</form>

  <% end if %>
<h4>Countries we ship to 
  <span>
    <a class="text-success ml-5 small small" href="countries-manage.asp?Add=yes" class="HomePageLinks">
      <i class="fa fa-plus-circle mr-1"></i>Add New
    </a>
  </span>
</h4>
<table class="table table-striped table-borderless table-hover w-50">
  <thead>
  <tr class="thead-dark">
    <th></th>
    <th>Display to public at checkout & accounts</th>
    <th>Country of origin option for admin products</th>
  </tr>
</thead>
  <% 
While ((Repeat1__numRows <> 0) AND (NOT rsGetCountries.EOF)) 
%>
    <tr>
      <td>
          <a class="btn btn-sm btn-danger mr-4" href="countries-manage.asp?Delete=yes&ID=<%=(rsGetCountries.Fields.Item("CountryID").Value)%>"><i class="fa fa-trash-alt"></i></a>
          <a href="countries-manage.asp?Edit=yes&ID=<%=(rsGetCountries.Fields.Item("CountryID").Value)%>" target="_top" class="LeftNavLinks"><%=(rsGetCountries.Fields.Item("Country").Value)%></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
      <td>
        <% If (rsGetCountries.Fields.Item("Display").Value) = 0 Then %>
        No
        <% else %>
        <span class="alert alert-success px-2 py-1 font-weight-bold">Yes</span>
      <% end if %></td>
      <td>
        <% If (rsGetCountries.Fields.Item("origin_toggle").Value) = 0 Then %>
        No
        <% else %>
        <span class="alert alert-info px-2 py-1 font-weight-bold">Yes</span>
      <% end if %></td>
    </tr>
    <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsGetCountries.MoveNext()
Wend
%>
</table>
</div>
</body>
</html>
<%
rsGetCountries.Close()
Set rsGetCountries = Nothing
%>
