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
  MM_editTable = "dbo.TBL_Companies"
  MM_editColumn = "companyID"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "add_company.asp"

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
Dim rsGetCompanies
Dim rsGetCompanies_numRows

Set rsGetCompanies = Server.CreateObject("ADODB.Recordset")
rsGetCompanies.ActiveConnection = MM_bodyartforms_sql_STRING
rsGetCompanies.Source = "SELECT * FROM dbo.TBL_Companies WHERE type = 'jewelry' AND display_AddEdit = 'yes' ORDER BY name ASC"
rsGetCompanies.CursorLocation = 3 'adUseClient
rsGetCompanies.LockType = 1 'Read-only records
rsGetCompanies.Open()

rsGetCompanies_numRows = 0
%>
<%
Dim rsInactive
Dim rsInactive_numRows

Set rsInactive = Server.CreateObject("ADODB.Recordset")
rsInactive.ActiveConnection = MM_bodyartforms_sql_STRING
rsInactive.Source = "SELECT * FROM dbo.TBL_Companies WHERE display_AddEdit <> 'yes' AND type = 'jewelry' ORDER BY name ASC"
rsInactive.CursorLocation = 3 'adUseClient
rsInactive.LockType = 1 'Read-only records
rsInactive.Open()

rsInactive_numRows = 0
%>
<html>
<head>
<title>List/add new jewelry company</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body>
<!--#include file="admin_header.asp"-->
<div class="p-3">
<% if request.querystring("Add") = "yes" then %>
<h4>Add a new vendor</h4>
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="FRM_AddCompany" id="FRM_AddCompany">
<table class="table table-sm table-striped table-borderless w-50">
    <tr> 
      <td style="width:25%">Company</td>
      <td> <input class="form-control form-control-sm" name="company" type="text"></td>
    </tr>
    <tr> 
      <td>Searchable brand name</td>
      <td> <input class="form-control form-control-sm" name="searchable_brand_tags" type="text"></td>
    </tr>
    <tr> 
      <td>Contact</td>
      <td><input class="form-control form-control-sm" name="contact" type="text"></td>
    </tr>
    <tr> 
      <td>Website</td>
      <td> <input class="form-control form-control-sm" name="website" type="text" ></td>
    </tr>
    <tr>
      <td>Notes</td>
      <td><textarea class="form-control form-control-sm" name="websiteLogin" cols="40" rows="3"></textarea></td>
    </tr>
    <tr> 
      <td>Email</td>
      <td> <input class="form-control form-control-sm" name="email" type="text"></td>
    </tr>
    <tr> 
      <td>Phone</td>
      <td><input class="form-control form-control-sm" name="phone" type="text"></td>
    </tr>
    <tr> 
      <td>Fax</td>
      <td> <input class="form-control form-control-sm" name="fax" type="text"></td>
    </tr>
    <tr>
      <td>Address</td>
      <td><textarea class="form-control form-control-sm" name="address" cols="40" rows="3"></textarea></td>
    </tr>
    <tr>
      <td>Type</td>
      <td><input name="type" type="radio" value="jewelry" checked>
        Jewlery
        <input class="ml-5" name="type" type="radio" value="beauty">
        Beauty</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>Display on Edit/Add page &nbsp;&nbsp;&nbsp;
       
        <input name="displayAdd" type="radio" value="yes" checked>
          yes 
          <input class="ml-5" name="displayAdd" type="radio" value="no">
        no </td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>Display on header dropdown &nbsp;&nbsp;&nbsp;
        <input name="displayHeader" type="radio" value="yes" checked>
yes
<input class="ml-5" name="displayHeader" type="radio" value="no">
no </td>
    </tr>
    
    <tr>
      <td>Display text logo in front of title</td>
      <td>
        <input name="ShowTextLogo" type="radio" value="Y">
yes
<input class="ml-5" name="ShowTextLogo" type="radio" value="N" checked>
no</td>
    </tr>
    <tr>
      <td>Product page logo</td>
      <td><input class="form-control form-control-sm" name="ProductLogo" type="text" > 
        (only the image name)</td>
    </tr>
    <tr>
      <td>Detail page logo</td>
      <td><input class="form-control form-control-sm" name="DetailLogo" type="text" >
        (only the image name)</td>
    </tr>
    <tr>
      <td>Custom order timeframe</td>
      <td><input class="form-control form-control-sm" name="preorder_timeframes" type="text">
	  
	  Custom order company? <input name="preorder_status" type="checkbox" value="1">
	  
        </td>
    </tr>
    <tr>
      <td>Vendor page FULL link</td>
      <td><input class="form-control form-control-sm" name="page" type="text" maxlength="250"></td>
    </tr>
    <tr>
      <td>SOP link</td>
      <td><input class="form-control form-control-sm" name="sop_link" type="text"></td>
    </tr>	
    <tr>
      <td>Weeks until receipt of order</td>
      <td>
		<select class="form-control form-control-sm w-auto" name="order_timeframes">
			<option value="0" selected>Weeks:</option>
			<%For i = 1 to 52%>
				<option value="<%=i%>"><%=i%> week<%If i>1 Then%>s<%End If%></option>
			<%Next%>
		</select>	  

	  </td>
    </tr>	
    <tr>
      <td>&nbsp;</td>
      <td><button class="btn btn-purple" type="submit" name="Submit">Add new vendor</button>
      <input type="hidden" name="dateadded" value="<%= date() %>"></td>
    </tr>
  </table>

<input type="hidden" name="MM_insert" value="FRM_AddCompany">
</form>
<%
' *** Insert Record: set variables

If (CStr(Request("MM_insert")) = "FRM_AddCompany") Then

  MM_editConnection = MM_bodyartforms_sql_STRING
  MM_editTable = "dbo.TBL_Companies"
  MM_editRedirectUrl = "add_company.asp"
  MM_fieldsStr  = "company|value|contact|value|website|value|websiteLogin|value|email|value|phone|value|fax|value|address|value|type|value|displayAdd|value|dateadded|value|displayHeader|value|ShowTextLogo|value|ProductLogo|value|DetailLogo|value|preorder_timeframes|value|page|value|sop_link|value|preorder_status|value|order_timeframes|value|searchable_brand_tags|value"
  MM_columnsStr = "name|',none,''|contact|',none,''|website|',none,''|website_login|',none,''|email|',none,''|phone|',none,''|fax|',none,''|address|',none,''|type|',none,''|display_AddEdit|',none,''|date_added|',none,NULL|display_Header|',none,''|ShowTextLogo|',none,''|ProductLogo|',none,''|DetailLogo|',none,''|preorder_timeframes|',none,''|page|',none,''|sop_link|',none,''|preorder_status|',none,''|order_timeframes|',none,''|searchable_brand_tags|',none,''"

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
rsEditCompany.Source = "SELECT * FROM dbo.TBL_Companies WHERE companyID = " + Replace(rsEditCompany__MMColParam, "'", "''") + ""
rsEditCompany.CursorLocation = 3 'adUseClient
rsEditCompany.LockType = 1 'Read-only records
rsEditCompany.Open()

rsEditCompany_numRows = 0
%>
<h4>Update a vendor</h4>
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="FRM_EditCompany" id="FRM_EditCompany">
<table class="table w-50 table-borderless table-striped">
    <tr> 
      <td style="width:25%">Company</td>
      <td> <input class="form-control form-control-sm" name="company" type="text" value="<%=(rsEditCompany.Fields.Item("name").Value)%>"></td>
    </tr>
    <tr> 
      <td>Searchable brand name</td>
      <td> <input class="form-control form-control-sm" name="searchable_brand_tags" type="text" value="<%= rsEditCompany.Fields.Item("searchable_brand_tags").Value %>"></td>
    </tr>
    <tr> 
      <td>Contact</td>
      <td><input class="form-control form-control-sm" name="contact" type="text" value="<%=(rsEditCompany.Fields.Item("contact").Value)%>"></td>
    </tr>
    <tr> 
      <td>Website</td>
      <td> <input class="form-control form-control-sm" name="website" type="text" value="<%=(rsEditCompany.Fields.Item("website").Value)%>"></td>
    </tr>
    <tr>
      <td>Notes</td>
      <td><textarea class="form-control form-control-sm" name="websiteLogin" cols="40" rows="3"><%=(rsEditCompany.Fields.Item("website_login").Value)%></textarea></td>
    </tr>
    <tr> 
      <td>Email</td>
      <td><input class="form-control form-control-sm" name="email" type="text" value="<%=(rsEditCompany.Fields.Item("email").Value)%>"></td>
    </tr>
    <tr> 
      <td>Phone</td>
      <td><input class="form-control form-control-sm" name="phone" type="text" value="<%=(rsEditCompany.Fields.Item("phone").Value)%>"></td>
    </tr>
    <tr> 
      <td>Fax</td>
      <td><input class="form-control form-control-sm" name="fax" type="text" value="<%=(rsEditCompany.Fields.Item("fax").Value)%>"></td>
    </tr>
    <tr>
      <td>Address</td>
      <td><textarea class="form-control form-control-sm" name="address" cols="40" rows="3"><%=(rsEditCompany.Fields.Item("address").Value)%></textarea></td>
    </tr>
    <tr>
      <td>Type</td>
      <td><input name="type" type="radio" value="jewelry" <% if (rsEditCompany.Fields.Item("type").Value) = "jewelry" then %>checked<% end if%>>
        Jewlery &nbsp;&nbsp;
        <input name="type" type="radio" value="beauty" <% if (rsEditCompany.Fields.Item("type").Value) = "beauty" then %>checked<% end if%>>
        Beauty</td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>Display on Edit/Add page &nbsp;&nbsp;&nbsp;
        <input name="displayAdd" type="radio" value="yes"  <% if (rsEditCompany.Fields.Item("display_AddEdit").Value) = "yes" then %> checked<% end if%>>
        yes
  <input class="ml-4" name="displayAdd" type="radio" value="no" <% if (rsEditCompany.Fields.Item("display_AddEdit").Value) <> "yes" then %> checked<% end if%>>
        no 
        
      </td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>Display on header dropdown &nbsp;
        <input name="displayHeader" type="radio" value="yes"  <% if (rsEditCompany.Fields.Item("display_Header").Value) = "yes" then %> checked<% end if%>>
        yes
  <input class="ml-4" name="displayHeader" type="radio" value="no"   <% if (rsEditCompany.Fields.Item("display_Header").Value) <> "yes" then %> checked<% end if%>>
        no 
        
      </td>
    </tr>
    <tr>
      <td>Display logo on site?</td>
      <td>
        <input name="ShowTextLogo" type="radio" value="Y" <% if (rsEditCompany.Fields.Item("ShowTextLogo").Value) = "Y" then %> checked<% end if%>>
        yes
        <input class="ml-4" name="ShowTextLogo" type="radio" value="N" <% if (rsEditCompany.Fields.Item("ShowTextLogo").Value) = "N" then %> checked<% end if%>>
        no</td>
    </tr>
    <tr>
      <td>Product page logo</td>
      <td><input class="form-control form-control-sm" name="ProductLogo" type="text" value="<%=(rsEditCompany.Fields.Item("ProductLogo").Value)%>">
          (only the image name)</td>
    </tr>
    <tr>
      <td>Detail page logo</td>
      <td><input class="form-control form-control-sm" name="DetailLogo" type="text"  value="<%=(rsEditCompany.Fields.Item("DetailLogo").Value)%>">
          (only the image name)</td>
    </tr>
    <tr>
      <td>Custom order timeframe</td>
      <td><input class="form-control form-control-sm mr-5" name="preorder_timeframes" type="text" value="<%=(rsEditCompany.Fields.Item("preorder_timeframes").Value)%>">
	 
	  Custom order company? <input name="preorder_status" type="checkbox" value="1" <% if rsEditCompany.Fields.Item("preorder_status").Value = 1 then %>checked<% end if %>>
        </td>
    <tr>
      <td>Vendor page FULL link</td>
      <td><input class="form-control form-control-sm" name="page" type="text" value="<%=(rsEditCompany.Fields.Item("page").Value)%>" maxlength="250"></td>
    </tr>
    <tr>
      <td>SOP link</td>
      <td><input class="form-control form-control-sm" name="sop_link" type="text" value="<%=rsEditCompany("sop_link")%>"></td>
    </tr>	
    <tr>
      <td>Weeks until receipt of order</td>
      <td>
		<select class="form-control form-control-sm w-auto" name="order_timeframes">
			<option value="0" <%If rsEditCompany("order_timeframes") = 0 Then%>selected<%End If%>>Weeks:</option>
			<%For i = 1 to 52%>
				<option value="<%=i%>" <%If rsEditCompany("order_timeframes") = i Then%>selected<%End If%>><%=i%> week<%If i>1 Then%>s<%End If%></option>
			<%Next%>
		</select>	  
	  </td>
    </tr>	
    <tr>
      <td>&nbsp;</td>
      <td><button class="btn btn-purple" type="submit" name="Submit">Update vendor</button></td>
    </tr>
  </table>

<input type="hidden" name="MM_update" value="FRM_EditCompany">
<input type="hidden" name="MM_recordId" value="<%= rsEditCompany.Fields.Item("companyID").Value %>">
</form>
<%
' *** Update Record: set variables

If (CStr(Request("MM_update")) = "FRM_EditCompany" And CStr(Request("MM_recordId")) <> "") Then

  MM_editConnection = MM_bodyartforms_sql_STRING
  MM_editTable = "dbo.TBL_Companies"
  MM_editColumn = "companyID"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "add_company.asp"
  MM_fieldsStr  = "company|value|contact|value|website|value|websiteLogin|value|email|value|phone|value|fax|value|address|value|type|value|displayAdd|value|displayHeader|value|ShowTextLogo|value|ProductLogo|value|DetailLogo|value|preorder_timeframes|value|page|value|sop_link|value|preorder_status|value|order_timeframes|value|searchable_brand_tags|value"
  MM_columnsStr = "name|',none,''|contact|',none,''|website|',none,''|website_login|',none,''|email|',none,''|phone|',none,''|fax|',none,''|address|',none,''|type|',none,''|display_AddEdit|',none,''|display_Header|',none,''|ShowTextLogo|',none,''|ProductLogo|',none,''|DetailLogo|',none,''|preorder_timeframes|',none,''|page|',none,''|sop_link|',none,''|preorder_status|',none,''|order_timeframes|',none,''|searchable_brand_tags|',none,''"
  

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
<form class="alert alert-danger" ACTION="<%=MM_editAction%>" METHOD="POST" name="FRM_DeleteCompany" id="FRM_DeleteCompany">
  <h5>Confirm deletion</h5>
  <button class="btn btn-sm btn-danger" type="submit" name="Submit2">Delete company</button>

    <input type="hidden" name="MM_delete" value="FRM_DeleteCompany">
    <input type="hidden" name="MM_recordId" value="<%= request.querystring("ID") %>">
</form>
<% end if %>
<p>
<table class="table table-sm table-striped table-borderless table-hover">
<thead class="thead-dark">
  <tr>
    <th colspan="7">ACTIVE jewelry companies </th>
  </tr>
  <tr>
    <th width="15%">Company <a class="btn btn-sm btn-info ml-3" href="add_company.asp?Add=yes"><i class="fa fa-plus-circle fa-lg"></i> Add new</a></th>
    <th width="15%">Searchable name</th>
	<th width="15%">Weeks until receipt of order</th>
    <th width="15%">Contact</th>
    <th width="10%">Phone</th>
    <th width="15%">SOP Link</th>	
    <th width="15%">Website</th>
  </tr>
 </thead>
  <% 
While NOT rsGetCompanies.EOF
%>
    <tr id="<%=(rsGetCompanies.Fields.Item("name").Value)%>">
      <td><a href="add_company.asp?Delete=yes&ID=<%=(rsGetCompanies.Fields.Item("companyID").Value)%>"><i class="fa fa-times-circle fa-lg text-danger mr-4"></i></a>
        <a href="add_company.asp?Edit=yes&ID=<%=(rsGetCompanies.Fields.Item("companyID").Value)%>" target="_top" class="LeftNavLinks"><%=(rsGetCompanies.Fields.Item("name").Value)%></a></td>
      <td><%=(rsGetCompanies.Fields.Item("searchable_brand_tags").Value)%></td>
	  <td><%=(rsGetCompanies.Fields.Item("order_timeframes").Value)%> weeks</td>
      <td><a href="mailto:<%=(rsGetCompanies.Fields.Item("email").Value)%>"><%=(rsGetCompanies.Fields.Item("contact").Value)%></a></td>
      <td><%=(rsGetCompanies.Fields.Item("phone").Value)%></td>
	  <td>
      <% if rsGetCompanies("sop_link") <> "" then %>
      <a href="<%=rsGetCompanies("sop_link")%>" title="<%=rsGetCompanies("sop_link")%>" target="_blank">View SOP</a>
      <% end if %>
    </td>
      <td><% if (rsGetCompanies.Fields.Item("website").Value) <> "" then %><a href="<%=(rsGetCompanies.Fields.Item("website").Value)%>" target="_blank">Visit website</a><% end if %></td>
    </tr>
    <% 
  rsGetCompanies.MoveNext()
Wend
%>
</table>

<table class="table table-sm table-striped table-borderless table-hover mt-5">
  <thead class="thead-dark">
  <tr>
    <th colspan="7">Inactive jewelry  companies </th>
  </tr>
  <tr >
    <th width="15%">Company</th>
    <th width="15%">Searchable name</th>
	<th width="15%">Weeks until receipt of order</th>
    <th width="15%">Contact</th>
    <th width="10%">Phone</th>
    <th width="15%">SOP Link</th>	
    <th width="15%">Website</th>
  </tr>
 </thead>
  <% 
While NOT rsInactive.EOF 
%>
    <tr>
      <td ><a class="text-danger" href="add_company.asp?Delete=yes&ID=<%=(rsInactive.Fields.Item("companyID").Value)%>"><i class="fa fa-times-circle fa-lg"></i></a>&nbsp;&nbsp;<a href="add_company.asp?Edit=yes&ID=<%=(rsInactive.Fields.Item("companyID").Value)%>" target="_top"><%=(rsInactive.Fields.Item("name").Value)%></a></td>
      <td><%=(rsInactive.Fields.Item("searchable_brand_tags").Value)%></td>
	  <td><%=(rsInactive.Fields.Item("order_timeframes").Value)%> weeks</td>
      <td><a href="mailto:<%=(rsInactive.Fields.Item("email").Value)%>"><%=(rsInactive.Fields.Item("contact").Value)%></a></td>
      <td><%=(rsInactive.Fields.Item("phone").Value)%></td>
	  <td><a target="blank_" href="<%=rsInactive("sop_link")%>" title="<%=rsInactive("sop_link")%>"><%=Left(rsInactive("sop_link"), 70)%><%If Len(rsInactive("sop_link")) > 70 Then%>...<%End If%></a></td>
      <td><% if (rsInactive.Fields.Item("website").Value) <> "" then %>
        <a href="<%=(rsInactive.Fields.Item("website").Value)%>" target="_blank">Visit website</a>
        <% end if %></td>
    </tr>
    <% 
  rsInactive.MoveNext()
Wend
%>

</table>


</div>
</body>
</html>
<%
rsGetCompanies.Close()
Set rsGetCompanies = Nothing

rsInactive.Close()
Set rsInactive = Nothing
%>
