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

if (CStr(Request("MM_delete")) = "FRM_DeleteCoupon" And CStr(Request("MM_recordId")) <> "") Then

  MM_editConnection = MM_bodyartforms_sql_STRING
  MM_editTable = "dbo.TBLDiscounts"
  MM_editColumn = "DiscountID"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "coupons_manage.asp"

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
Dim rsGetCoupons__MMColParam
rsGetCoupons__MMColParam = "A"
If (Request("MM_EmptyValue") <> "") Then 
  rsGetCoupons__MMColParam = Request("MM_EmptyValue")
End If
%>
<%
Dim rsGetCoupons
Dim rsGetCoupons_numRows

Set rsGetCoupons = Server.CreateObject("ADODB.Recordset")
rsGetCoupons.ActiveConnection = MM_bodyartforms_sql_STRING
rsGetCoupons.Source = "SELECT * FROM dbo.TBLDiscounts WHERE DateExpired >= '" & now() & "' AND coupon_single_use = 0 ORDER BY DiscountID ASC"
rsGetCoupons.CursorLocation = 3 'adUseClient
rsGetCoupons.LockType = 1 'Read-only records
rsGetCoupons.Open()

rsGetCoupons_numRows = 0
%>
<%
Dim rsInactive__MMColParam
rsInactive__MMColParam = "A"
If (Request("MM_EmptyValue") <> "") Then 
  rsInactive__MMColParam = Request("MM_EmptyValue")
End If
%>
<%
Dim rsInactive
Dim rsInactive_numRows

Set rsInactive = Server.CreateObject("ADODB.Recordset")
rsInactive.ActiveConnection = MM_bodyartforms_sql_STRING
rsInactive.Source = "SELECT * FROM dbo.TBLDiscounts WHERE DateExpired < '" & now() & "' AND DateExpired > '" & now() - 720 & "' AND coupon_single_use = 0 ORDER BY DateExpired DESC"
rsInactive.CursorLocation = 3 'adUseClient
rsInactive.LockType = 1 'Read-only records
rsInactive.Open()

rsInactive_numRows = 0
%>
<%
Dim rsEditCompany__MMColParam
rsEditCompany__MMColParam = "1"
If (Request.QueryString("ID") <> "") Then 
  rsEditCompany__MMColParam = Request.QueryString("ID")
End If
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
rsGetCoupons_numRows = rsGetCoupons_numRows + Repeat1__numRows
%>
<%
Dim Repeat2__numRows
Dim Repeat2__index

Repeat2__numRows = -1
Repeat2__index = 0
rsGetBeauty_numRows = rsGetBeauty_numRows + Repeat2__numRows
%>
<%
Dim Repeat3__numRows
Dim Repeat3__index

Repeat3__numRows = -1
Repeat3__index = 0
rsInactive_numRows = rsInactive_numRows + Repeat3__numRows
%>
<%
Dim rsGetCompany
Dim rsGetCompany_numRows

Set rsGetCompany = Server.CreateObject("ADODB.Recordset")
rsGetCompany.ActiveConnection = MM_bodyartforms_sql_STRING
rsGetCompany.Source = "SELECT companyID, name, display_AddEdit FROM dbo.TBL_Companies WHERE display_AddEdit = 'yes' AND type = 'jewelry' ORDER BY name ASC"
rsGetCompany.CursorLocation = 3 'adUseClient
rsGetCompany.LockType = 1 'Read-only records
rsGetCompany.Open()

rsGetCompany_numRows = 0
%>

<%
Dim Repeat4__numRows
Dim Repeat4__index

Repeat4__numRows = -1
Repeat4__index = 0
rsGetCompany_numRows = rsGetCompany_numRows + Repeat4__numRows
%>

<html>
<head>
<title>Manage coupons</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body>
<!--#include file="admin_header.asp"-->
<div  class="p-3">
<% if request.querystring("Add") = "yes" then %>

<form ACTION="<%=MM_editAction%>" METHOD="POST" name="FRM_AddCoupon" id="FRM_AddCoupon">
<table class="table table-striped table-borderless table-hover">
    <tr> 
      <td >Description</td>
      <td> <input name="description" type="text" id="description" size="30"></td>
    </tr>
    <tr > 
      <td>CODE (that customers use) &nbsp;</td>
      <td ><input name="code" type="text" id="code" size="30"></td>
    </tr>
    <tr > 
      <td>Type of discount &nbsp;</td>
      <td ><select name="Type" id="Type">
        <option value="Percentage" selected>Percentage</option>
        <option value="Fixed">Fixed</option>
      </select></td>
    </tr>
    <tr > 
      <td>&nbsp;</td>
      <td > %
<input name="percentage" type="text" id="percentage" value="0" size="4">
        &nbsp;&nbsp;&nbsp;Fixed
        <input name="fixed" type="text" id="fixed" value="0" size="4"></td>
    </tr>
    <tr >
      <td>Start date &nbsp;</td>
      <td ><input name="DateActive" type="text" id="DateActive" value="<%= date() %>" size="20"></td>
    </tr>
    <tr > 
      <td>End date &nbsp;</td>
      <td ><input name="expdate" type="text" id="expdate" value="<%= date() + 3 %> 11:59:00 PM" size="20"> 
        (mm/dd/yyyy) </td>
    </tr>
    <tr >
      <td>Active</td>
      <td ><input name="active" type="radio" value="A" checked>
        YES &nbsp;&nbsp;
        <input name="active" type="radio" value="N">
        No</td>
    </tr>
    
    <tr >
      <td  valign="top">Brand</td>
      <td valign="top"><select name="BrandName" id="BrandName">
<option value="None" selected>None</option>
                <% 
While ((Repeat4__numRows <> 0) AND (NOT rsGetCompany.EOF)) 
%>
                <option value="<%=(rsGetCompany.Fields.Item("name").Value)%>"><%=(rsGetCompany.Fields.Item("name").Value)%></option>
                
                <% 
  Repeat4__index=Repeat4__index+1
  Repeat4__numRows=Repeat4__numRows-1
  rsGetCompany.MoveNext()
Wend
rsGetCompany.Requery()
%> 
      </select></td>
    </tr>
    <tr >
      <td  valign="top">Clearance?</td>
      <td valign="top"><select name="Clearance" id="Clearance">
                <option value="None" selected>None</option>
                <option value="Clearance">Clearance</option>
                <option value="limited">Limited</option>
                <option value="Discontinued">Discontinued</option>
                <option value="One time buy">One time buy</option>
                <option value="OneDay">One day sale</option>
      </select> 
      -- DOES NOT WORK!!</td>
    </tr>
    <tr >
      <td  valign="top">DOES NOT WORK</td>
      <td valign="top"><p>
        
          Exclude clearance/limited/sale items from discount?<br>
          <input name="ExcludeSaleItems" type="radio" id="ExcludeSaleItems" value="1" checked>
          Yes
          &nbsp;&nbsp;&nbsp;
          <input type="radio" name="ExcludeSaleItems" value="0" id="ExcludeSaleItems">
          No
        <br>
      </p>
	  
	  
	  	Show on website: 
<input type="radio" name="show_on_website" value="1">
        Yes
      &nbsp;&nbsp;&nbsp;
        <input type="radio" name="show_on_website" value="0" checked>No
<br/>
Website text: <br/>
<textarea maxlength="250" name="website_text"></textarea>
<input type="hidden" name="MM_update" value="FRM_EditCoupon">
<input type="hidden" name="MM_recordId" value="<%= request.querystring("ID") %>">
	  </td>
    </tr>
    <tr >
      <td  valign="top">&nbsp;</td>
      <td valign="top"><input type="submit" name="Submit" value="Submit" class="submit_button">
      <input type="hidden" name="dateadded" value="<%= now() %>"></td>
    </tr>
  </table>


<input type="hidden" name="MM_insert" value="FRM_AddCoupon">
</form>
<%
' *** Insert Record: set variables

If (CStr(Request("MM_insert")) = "FRM_AddCoupon") Then

  MM_editConnection = MM_bodyartforms_sql_STRING
  MM_editTable = "dbo.TBLDiscounts"
  MM_editRedirectUrl = "coupons_manage.asp"
  MM_fieldsStr  = "description|value|code|value|Type|value|percentage|value|fixed|value|expdate|value|active|value|dateadded|value|DateActive|value|BrandName|value|Clearance|value|ExcludeSaleItems|value|show_on_website|value|website_text|value"
  MM_columnsStr = "DiscountDescription|',none,''|DiscountCode|',none,''|DiscountType|',none,''|DiscountPercent|none,none,NULL|DiscountPrice|none,none,NULL|DateExpired|',none,NULL|Active|',none,''|DateAdded|',none,NULL|DateActive|',none,NULL|BrandName|',none,NULL|Clearance|',none,NULL|ExcludeSaleItems|',none,NULL|show_on_website|',none,''|website_text|',none,''"

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
Dim rsEditCompany
Dim rsEditCompany_numRows

Set rsEditCompany = Server.CreateObject("ADODB.Recordset")
rsEditCompany.ActiveConnection = MM_bodyartforms_sql_STRING
rsEditCompany.Source = "SELECT * FROM dbo.TBLDiscounts WHERE DiscountID = " + Replace(rsEditCompany__MMColParam, "'", "''") + ""
rsEditCompany.CursorLocation = 3 'adUseClient
rsEditCompany.LockType = 1 'Read-only records
rsEditCompany.Open()

rsEditCompany_numRows = 0
%>
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="FRM_EditCoupon" id="FRM_EditCoupon" class="basic_form">
&nbsp;<br>
<table class="admin-table">
  <tr >
    <td>Description&nbsp;</td>
    <td ><input name="description" type="text" id="description" value="<%=(rsEditCompany.Fields.Item("DiscountDescription").Value)%>" size="30"></td>
  </tr>
  <tr >
    <td>CODE (that customers use) &nbsp;</td>
    <td ><input name="code" type="text" id="code" value="<%=(rsEditCompany.Fields.Item("DiscountCode").Value)%>" size="30"></td>
  </tr>
  <tr >
    <td>Type of discount &nbsp;</td>
    <td ><select name="type" id="type">
      <option value="<%=(rsEditCompany.Fields.Item("DiscountType").Value)%>" selected><%=(rsEditCompany.Fields.Item("DiscountType").Value)%></option>
	  <option value="Percentage">Percentage</option>
      <option value="Fixed">Fixed</option>
    </select></td>
  </tr>
  <tr >
    <td>&nbsp;</td>
    <td > %
      <input name="percentage" type="text" id="percentage" value="<%=(rsEditCompany.Fields.Item("DiscountPercent").Value)%>" size="4">
      &nbsp;&nbsp;&nbsp;Fixed
      <input name="fixed" type="text" id="fixed" value="<%=(rsEditCompany.Fields.Item("DiscountPrice").Value)%>" size="4"></td>
  </tr>
  <tr >
    <td>Start date&nbsp;</td>
    <td ><input name="DateActive" type="text" id="DateActive" value="<%=(rsEditCompany.Fields.Item("DateActive").Value)%>" size="20"></td>
  </tr>
  <tr >
    <td>End date &nbsp;</td>
    <td ><input name="expdate" type="text" id="expdate" value="<%=(rsEditCompany.Fields.Item("DateExpired").Value)%>" size="20">
      (mm/dd/yyyy) </td>
  </tr>
  <tr >
    <td>Active</td>
    <td ><input name="active" type="radio" value="A" <% if (rsEditCompany.Fields.Item("Active").Value) = "A" then%>checked<%end if %>>
      YES &nbsp;&nbsp;
      <input name="active" type="radio" value="N" <% if (rsEditCompany.Fields.Item("Active").Value) <> "A" then%>checked<%end if %>>
      No</td>
  </tr>
  <tr >
    <td  valign="top">Brand</td>
    <td valign="top"><select name="BrandName" id="BrandName">
      <option value="None" >None</option>
<option value="<%=(rsEditCompany.Fields.Item("BrandName").Value)%>" selected><%=(rsEditCompany.Fields.Item("BrandName").Value)%></option>
<option>-------------------------</option>
      <% 
While ((Repeat4__numRows <> 0) AND (NOT rsGetCompany.EOF)) 
%>
      <option value="<%=(rsGetCompany.Fields.Item("name").Value)%>"><%=(rsGetCompany.Fields.Item("name").Value)%></option>
      <% 
  Repeat4__index=Repeat4__index+1
  Repeat4__numRows=Repeat4__numRows-1
  rsGetCompany.MoveNext()
Wend
%>
    </select></td>
  </tr>
  <tr >
    <td  valign="top">Clearance?</td>
    <td valign="top"><select name="Clearance" id="Clearance">
  <option value="<%=(rsEditCompany.Fields.Item("Clearance").Value)%>" selected><%=(rsEditCompany.Fields.Item("Clearance").Value)%></option>
  <option>-------------------------</option>
  <option value="None">None</option>
      <option value="Clearance">Clearance</option>
      <option value="limited">Limited</option>
      <option value="Discontinued">Discontinued</option>
      <option value="One time buy">One time buy</option>
      <option value="OneDay">One day sale</option>
      </select>
    DOES NOT WORK</td>
  </tr>
  <tr >
    <td  valign="top">DOES NOT WORK</td>
    <td valign="top"><p>
      Exclude clearance/limited/sale items from discount?<br>
        <input name="ExcludeSaleItems" type="radio" id="ExcludeSaleItems" value="1"<% if rsEditCompany.Fields.Item("ExcludeSaleItems").Value = 1 then %> checked<% end if %>>
        Yes
      &nbsp;&nbsp;&nbsp;
        <input type="radio" name="ExcludeSaleItems" value="0" id="ExcludeSaleItems" <% if rsEditCompany.Fields.Item("ExcludeSaleItems").Value = 0 then %> checked<% end if %>>
        No
      <br>
    </p>
	Show on website: 
<input name="show_on_website" type="radio" value="1"<% if rsEditCompany.Fields.Item("show_on_website").Value = 1 then %> checked<% end if %>>
        Yes
      &nbsp;&nbsp;&nbsp;
        <input type="radio" name="show_on_website" value="0" <% if rsEditCompany.Fields.Item("show_on_website").Value = 0 then %> checked<% end if %>>No
<br/>
Website text: <br/>
<textarea maxlength="250" name="website_text"><%=(rsEditCompany.Fields.Item("website_text").Value)%></textarea>
<input type="hidden" name="MM_update" value="FRM_EditCoupon">
<input type="hidden" name="MM_recordId" value="<%= request.querystring("ID") %>">
	
	</td>
  </tr>
  <tr >
    <td  valign="top">&nbsp;</td>
    <td valign="top"><input type="submit" name="Submit3" value="Submit" class="submit_button"></td>
  </tr>
</table>
</form>
<%
' *** Update Record: set variables

If (CStr(Request("MM_update")) = "FRM_EditCoupon" And CStr(Request("MM_recordId")) <> "") Then

  MM_editConnection = MM_bodyartforms_sql_STRING
  MM_editTable = "dbo.TBLDiscounts"
  MM_editColumn = "DiscountID"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "coupons_manage.asp"
  MM_fieldsStr  = "description|value|code|value|type|value|percentage|value|fixed|value|expdate|value|active|value|DateActive|value|BrandName|value|Clearance|value|ExcludeSaleItems|value|show_on_website|value|website_text|value"
  MM_columnsStr = "DiscountDescription|',none,''|DiscountCode|',none,''|DiscountType|',none,''|DiscountPercent|none,none,NULL|DiscountPrice|none,none,NULL|DateExpired|',none,NULL|Active|',none,''|DateActive|',none,''|BrandName|',none,''|Clearance|',none,''|ExcludeSaleItems|',none,''|show_on_website|',none,''|website_text|',none,''"

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
<% end if %>

<table class="table table-striped table-borderless table-hover">
<thead class="thead-dark">
  <tr>
    <th colspan="6">ACTIVE COUPONS</td>
  </th>
  <tr>
    <th width="40%">Desription <a href="coupons_manage.asp?Add=yes"><i class="fa fa-plus-circle fa-lg" style="color:#088A08;margin:0 .5em 0 2em"></i>ADD NEW</a></th>
    <th width="20%">Code</th>
    <th width="10%">Type</th>
    <th width="10%">Begins</th>
    <th width="10%">Expires</th>
    <th width="10%">Display on Site</th>
  </tr>
</thead>
  <% 
While ((Repeat1__numRows <> 0) AND (NOT rsGetCoupons.EOF)) 

if rsGetCoupons.Fields.Item("show_on_website").Value = 1 then
	show_on_site = "<span class=""notice-eco"">Yes</span>"
else
	show_on_site = "No"
end if
%>
    <tr>
      <td><a href="coupons_manage.asp?Edit=yes&ID=<%=(rsGetCoupons.Fields.Item("DiscountID").Value) %>" target="_top"><%=(rsGetCoupons.Fields.Item("DiscountDescription").Value)%></a></td>
      <td><%=(rsGetCoupons.Fields.Item("DiscountCode").Value)%></td>
      <td><%=(rsGetCoupons.Fields.Item("DiscountType").Value)%></td>
      <td><%=(rsGetCoupons.Fields.Item("DateActive").Value)%></td>
      <td><%=(rsGetCoupons.Fields.Item("DateExpired").Value)%></td>
      <td><%= show_on_site %></td>
    </tr>
    <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsGetCoupons.MoveNext()
Wend
%>
</table>
<p>&nbsp;</p>
<table class="admin-table">
	<thead>
  <tr>
    <th colspan="4">INACTIVE COUPONS</th>
  </tr>
  <tr>
    <th width="40%">Desription</th>
    <th width="20%">Code</th>
    <th width="20%">Type</th>
    <th width="10%">Expires</th>
  </tr>
 </thead>
  <% 
While ((Repeat3__numRows <> 0) AND (NOT rsInactive.EOF)) 
%>
    <tr>
      <td><a href="coupons_manage.asp?Edit=yes&ID=<%=(rsInactive.Fields.Item("DiscountID").Value)%>" target="_top"><%=(rsInactive.Fields.Item("DiscountDescription").Value)%></a></td>
      <td><%=(rsInactive.Fields.Item("DiscountCode").Value)%></td>
      <td><%=(rsInactive.Fields.Item("DiscountType").Value)%></td>
      <td><%=(rsInactive.Fields.Item("DateExpired").Value)%></td>
    </tr>
    <% 
  Repeat3__index=Repeat3__index+1
  Repeat3__numRows=Repeat3__numRows-1
  rsInactive.MoveNext()
Wend
%>
</table>
</div>
</body>
</html>
<%
rsGetCoupons.Close()
Set rsGetCoupons = Nothing
%>
<%
rsInactive.Close()
Set rsInactive = Nothing
%>
<%
rsGetCompany.Close()
Set rsGetCompany = Nothing
%>
