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

If (CStr(Request("MM_update")) = "FRM_EditCoupon" And CStr(Request("MM_recordId")) <> "") Then

  MM_editConnection = MM_bodyartforms_sql_STRING
  MM_editTable = "dbo.TBL_ShippingMethods"
  MM_editColumn = "IDShipping"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "manage_shippingmethods.asp"
  MM_fieldsStr  = "ShippingName|value|ShippingAmount|value|ShippingDesc_Public|value|ShippingDesc_Private|value|ShippingWeight|value|ShippingType|value|ShippingDiscount|value|DiscountSubtotal|value|SortOrder|value|FAQShow|value|est_days_min|value|ShippingWeightMIN|value|est_days_max|value"
  MM_columnsStr = "ShippingName|',none,''|ShippingAmount|none,none,NULL|ShippingDesc_Public|',none,''|ShippingDesc_Private|',none,''|ShippingWeight|none,none,NULL|ShippingType|',none,''|ShippingDiscount|none,none,NULL|DiscountSubtotal|none,none,NULL|SortOrder|none,none,NULL|FAQShow|none,none,NULL|est_days_min|',none,''|ShippingWeightMIN|',none,''|est_days_max|',none,''"

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

If (CStr(Request("MM_insert")) = "FRM_AddCoupon") Then

  MM_editConnection = MM_bodyartforms_sql_STRING
  MM_editTable = "dbo.TBL_ShippingMethods"
  MM_editRedirectUrl = "manage_shippingmethods.asp"
  MM_fieldsStr  = "ShippingName|value|ShippingAmout|value|ShippingDesc_Public|value|ShippingDesc_Private|value|ShippingWeight|value|ShippingType|value|ShippingDiscount|value|DiscountSubtotal|value|item_order|value|FAQShow|value|est_days_min|value|ShippingWeightMIN|value|est_days_max|value"
  MM_columnsStr = "ShippingName|',none,''|ShippingAmount|none,none,NULL|ShippingDesc_Public|',none,''|ShippingDesc_Private|',none,''|ShippingWeight|none,none,NULL|ShippingType|',none,''|ShippingDiscount|none,none,NULL|DiscountSubtotal|none,none,NULL|SortOrder|none,none,NULL|FAQShow|none,none,NULL|est_days_min|',none,''|ShippingWeightMIN|',none,''|est_days_max|',none,''"

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

if (CStr(Request("MM_delete")) = "FRM_DeleteCoupon" And CStr(Request("MM_recordId")) <> "") Then

  MM_editConnection = MM_bodyartforms_sql_STRING
  MM_editTable = "dbo.TBL_ShippingMethods"
  MM_editColumn = "IDShipping"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "manage_shippingmethods.asp"

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
Dim rsGetShipping__MMColParam
rsGetShipping__MMColParam = "A"
If (Request("MM_EmptyValue") <> "") Then 
  rsGetShipping__MMColParam = Request("MM_EmptyValue")
End If
%>
<%
Dim rsGetShipping
Dim rsGetShipping_numRows

Set rsGetShipping = Server.CreateObject("ADODB.Recordset")
rsGetShipping.ActiveConnection = MM_bodyartforms_sql_STRING
rsGetShipping.Source = "SELECT * FROM dbo.TBL_ShippingMethods ORDER BY SortOrder ASC"
rsGetShipping.CursorLocation = 3 'adUseClient
rsGetShipping.LockType = 1 'Read-only records
rsGetShipping.Open()

rsGetShipping_numRows = 0
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
rsGetONE.Source = "SELECT * FROM dbo.TBL_ShippingMethods WHERE IDShipping = " + Replace(rsGetONE__MMColParam, "'", "''") + ""
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
rsGetShipping_numRows = rsGetShipping_numRows + Repeat1__numRows
%>
<html>
<head>
<title>Manage shipping methods</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body>
<!--#include file="admin_header.asp"-->
<div class="p-3">
<% if request.querystring("Add") = "yes" then %>
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="FRM_AddCoupon" id="FRM_AddCoupon">
&nbsp;<br>
<table width="50%" border="0" cellspacing="0" cellpadding="1">
    <tr align="left"> 
      <td align="right" valign="middle" class="pricegauge">Shipping name &nbsp;</td>
      <td valign="middle" class="pricegauge"> <input name="ShippingName" type="text" class="adminfields" id="ShippingName" size="30"></td>
    </tr>
    <tr align="left"> 
      <td align="right" valign="middle" class="pricegauge">Shipping amount&nbsp;&nbsp; </td>
      <td valign="middle" class="pricegauge"><input name="ShippingAmout" type="text" class="adminfields" id="ShippingAmout" size="7"></td>
    </tr>
	    <tr align="left">
      <td align="right" valign="top" class="pricegauge">Estimated delivery&nbsp;&nbsp;</td>
      <td valign="top" class="pricegauge">MIN days: <input name="est_days_min" type="text" class="adminfields" id="ShippingDiscount" value="0" size="6">&nbsp;&nbsp;&nbsp;&nbsp;MAX days: <input name="est_days_max" type="text" class="adminfields" id="ShippingDiscount" value="0" size="6">
      </td>
    </tr>
	<tr align="left">
      <td align="right" valign="top" class="pricegauge">Discounted price&nbsp;&nbsp;</td>
      <td valign="top" class="pricegauge"><input name="ShippingDiscount" type="text" class="adminfields" id="ShippingDiscount" value="0" size="6">
      &nbsp;&nbsp;&nbsp;Subtotal before discount? 
      <input name="DiscountSubtotal" type="text" class="adminfields" id="DiscountSubtotal" value="0" size="6"></td>
    </tr>
    <tr align="left"> 
      <td align="right" valign="middle" class="pricegauge">Shipping description (PRIVATE) &nbsp;</td>
      <td valign="middle" class="pricegauge"><textarea name="ShippingDesc_Private" cols="50" rows="3" class="adminfields" id="ShippingDesc_Private"></textarea></td>
    </tr>
    <tr align="left"> 
      <td align="right" valign="middle" class="pricegauge">Public description &nbsp;</td>
      <td valign="middle" class="pricegauge"><textarea name="ShippingDesc_Public" cols="50" rows="3" class="adminfields" id="ShippingDesc_Public"></textarea></td>
    </tr>
    <tr align="left">
      <td align="right" valign="middle" class="pricegauge">Weight restrictions (ounces) &nbsp;</td>
      <td valign="middle" class="pricegauge">MIN weight: <input name="ShippingWeightMIN" type="text" class="adminfields" id="ShippingWeightMIN" value="0" size="6">&nbsp;&nbsp;&nbsp;MAX weight: <input name="ShippingWeight" type="text" class="adminfields" id="ShippingWeight" value="0" size="6"></td>
    </tr>
    
    <tr align="left">
      <td align="right" valign="top" class="pricegauge">Shipping type&nbsp;&nbsp;&nbsp; </td>
      <td valign="top" class="pricegauge"><label>
        <select name="ShippingType" class="adminfields" id="ShippingType">
          <option value="USA">USA</option>
		  <option value="Canada">Canada</option>
          <option value="International">International</option>
        </select>
      </label></td>
    </tr>
    <tr align="left">
      <td align="right" valign="top" class="pricegauge">Sort order &nbsp;&nbsp;</td>
      <td valign="top" class="pricegauge"><select name="item_order" class="adminfields" id="item_order">
        <option value="0" selected>0</option>
        <option value="1">1</option>
        <option value="2">2</option>
        <option value="3">3</option>
        <option value="4">4</option>
        <option value="5">5</option>
        <option value="6">6</option>
        <option value="7">7</option>
        <option value="8">8</option>
        <option value="9">9</option>
        <option value="10">10</option>
        <option value="11">11</option>
        <option value="12">12</option>
        <option value="13">13</option>
        <option value="14">14</option>
        <option value="15">15</option>
        <option value="16">16</option>
        <option value="17">17</option>
        <option value="18">18</option>
        <option value="19">19</option>
        <option value="20">20</option>
            </select></td>
    </tr>
    <tr align="left">
      <td align="right" valign="top" class="pricegauge">Show on FAQ page? &nbsp;&nbsp;</td>
      <td valign="top" class="pricegauge"><select name="FAQShow" class="adminfields" id="FAQShow">
        <option value="1">Yes</option>
        <option value="0" selected>No</option>
            </select></td>
    </tr>
    <tr align="left">
      <td align="right" valign="top" class="pricegauge">&nbsp;</td>
      <td valign="top" class="pricegauge"><input type="submit" name="Submit" value="Submit"></td>
    </tr>
  </table>



<input type="hidden" name="MM_insert" value="FRM_AddCoupon">
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
rsEditCompany.Source = "SELECT * FROM dbo.TBL_ShippingMethods WHERE IDShipping = " + Replace(rsEditCompany__MMColParam, "'", "''") + ""
rsEditCompany.CursorLocation = 3 'adUseClient
rsEditCompany.LockType = 1 'Read-only records
rsEditCompany.Open()

rsEditCompany_numRows = 0
%>
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="FRM_EditCoupon" id="FRM_EditCoupon">
&nbsp;<br>
<table width="50%" border="0" cellspacing="0" cellpadding="1">
  <tr align="left">
    <td align="right" valign="middle" class="pricegauge">Shipping name &nbsp;</td>
    <td valign="middle" class="pricegauge"><input name="ShippingName" type="text" class="adminfields" id="ShippingName" value="<%=(rsGetONE.Fields.Item("ShippingName").Value)%>" size="30"></td>
  </tr>
  <tr align="left">
    <td align="right" valign="middle" class="pricegauge">Shipping amount&nbsp;&nbsp;</td>
    <td valign="middle" class="pricegauge"><input name="ShippingAmount" type="text" class="adminfields" id="ShippingAmount" value="<%=(rsGetONE.Fields.Item("ShippingAmount").Value)%>" size="7"></td>
  </tr>
    <tr align="left">
    <td align="right" valign="top" class="pricegauge">Discounted price&nbsp;&nbsp;</td>
    <td valign="top" class="pricegauge"><input name="ShippingDiscount" type="text" class="adminfields" id="ShippingDiscount" value="<%=(rsGetONE.Fields.Item("ShippingDiscount").Value)%>" size="6">
      &nbsp;&nbsp;&nbsp;Subtotal before discount?
      <input name="DiscountSubtotal" type="text" class="adminfields" id="DiscountSubtotal" value="<%=(rsGetONE.Fields.Item("DiscountSubtotal").Value)%>" size="6"></td>
  </tr>
  	    <tr align="left">
      <td align="right" valign="top" class="pricegauge">Estimated delivery&nbsp;&nbsp;</td>
      <td valign="top" class="pricegauge">MIN days: <input name="est_days_min" type="text" class="adminfields" id="ShippingDiscount" value="<%=(rsGetONE.Fields.Item("est_days_min").Value)%>" size="6">&nbsp;&nbsp;&nbsp;&nbsp;MAX days: <input name="est_days_max" type="text" class="adminfields" id="ShippingDiscount" value="<%=(rsGetONE.Fields.Item("est_days_max").Value)%>" size="6">
      </td>
    </tr>
  <tr align="left">
    <td align="right" valign="middle" class="pricegauge">Shipping description (PRIVATE) &nbsp;</td>
    <td valign="middle" class="pricegauge"><textarea name="ShippingDesc_Private" cols="50" rows="3" class="adminfields" id="ShippingDesc_Private"><%=(rsGetONE.Fields.Item("ShippingDesc_Private").Value)%></textarea></td>
  </tr>
  <tr align="left">
    <td align="right" valign="middle" class="pricegauge">Public description &nbsp;</td>
    <td valign="middle" class="pricegauge"><textarea name="ShippingDesc_Public" cols="50" rows="3" class="adminfields" id="ShippingDesc_Public"><%=(rsGetONE.Fields.Item("ShippingDesc_Public").Value)%></textarea></td>
  </tr>

  <tr align="left">
    <td align="right" valign="middle" class="pricegauge">Weight restrictions (ounces) &nbsp;</td>
    <td valign="middle" class="pricegauge">MIN weight: <input name="ShippingWeightMIN" type="text" class="adminfields" id="ShippingWeightMIN" value="<%=(rsGetONE.Fields.Item("ShippingWeightMIN").Value)%>" size="6">&nbsp;&nbsp;&nbsp;MAX weight: <input name="ShippingWeight" type="text" class="adminfields" id="ShippingWeight" value="<%=(rsGetONE.Fields.Item("ShippingWeight").Value)%>" size="6"></td>
  </tr>
  <tr align="left">
    <td align="right" valign="top" class="pricegauge">Shipping type&nbsp;&nbsp;&nbsp; </td>
    <td valign="top" class="pricegauge"><label>
      <select name="ShippingType" class="adminfields" id="ShippingType">
      <option value="<%=(rsGetONE.Fields.Item("ShippingType").Value)%>" selected><%=(rsGetONE.Fields.Item("ShippingType").Value)%></option>	          <option value="USA">USA</option>
		  <option value="Canada">Canada</option>
          <option value="International">International</option>
      </select>
    </label></td>
  </tr>
  <tr align="left">
    <td align="right" valign="top" class="pricegauge">Sort order &nbsp;&nbsp;</td>
    <td valign="top" class="pricegauge"><select name="SortOrder" class="adminfields" id="SortOrder">
 <option value="<%=(rsGetONE.Fields.Item("SortOrder").Value)%>" selected><%=(rsGetONE.Fields.Item("SortOrder").Value)%></option>
         <option value="0">0</option>
        <option value="1">1</option>
        <option value="2">2</option>
        <option value="3">3</option>
        <option value="4">4</option>
        <option value="5">5</option>
        <option value="6">6</option>
        <option value="7">7</option>
        <option value="8">8</option>
        <option value="9">9</option>
        <option value="10">10</option>
        <option value="11">11</option>
        <option value="12">12</option>
        <option value="13">13</option>
        <option value="14">14</option>
        <option value="15">15</option>
        <option value="16">16</option>
        <option value="17">17</option>
        <option value="18">18</option>
        <option value="19">19</option>
        <option value="20">20</option>
    </select></td>
  </tr>
  <tr align="left">
    <td align="right" valign="top" class="pricegauge">Show on FAQ page? &nbsp;&nbsp;</td>
    <td valign="top" class="pricegauge"><select name="FAQShow" class="adminfields" id="FAQShow">
 <option value="<%=(rsGetONE.Fields.Item("FAQShow").Value)%>" selected><%=(rsGetONE.Fields.Item("FAQShow").Value)%></option>        <option value="1">Yes</option>
        <option value="0">No</option>
    </select></td>
  </tr>
  
  <tr align="left">
    <td align="right" valign="top" class="pricegauge">&nbsp;</td>
    <td valign="top" class="pricegauge"><input type="submit" name="Submit3" value="Submit"></td>
  </tr>
</table>

<input type="hidden" name="MM_update" value="FRM_EditCoupon">
<input type="hidden" name="MM_recordId" value="<%= request.querystring("ID") %>">
</form>
<% end if %>
<% if request.querystring("Delete") = "yes" then %>
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="FRM_DeleteCoupon" id="FRM_DeleteCoupon">
  <div class="alert alert-danger">Confirm deletion <input class="btn btn-sm btn-danger ml-4" type="submit" name="Submit2" value="Delete">
  </div>
    <input type="hidden" name="MM_delete" value="FRM_DeleteCoupon">
<input type="hidden" name="MM_recordId" value="<%= request.querystring("ID") %>">
</form>
<% end if %>
<h5>Checkout Shipping Options</h5>
<table class="table table-striped table-borderless table-hover">
<thead class="thead-dark">
  <tr>
    <th width="15%">Name
      <a class="ml-3 text-success" href="manage_shippingmethods.asp?Add=yes"><i class="fa fa-plus-circle mr-1"></i>Add New</a></th>
    <th width="5%">Amount</th>
    <th width="30%">Description PUBLIC </th>
    <th width="25%">&nbsp;</th>
  </tr>
</thead>
  <% 
While ((Repeat1__numRows <> 0) AND (NOT rsGetShipping.EOF)) 
%>
    <tr>
      <td><strong><%=(rsGetShipping.Fields.Item("ShippingType").Value)%>
      </strong><br>
      <a class="btn btn-sm btn-danger mr-2" href="manage_shippingmethods.asp?Delete=yes&ID=<%=(rsGetShipping.Fields.Item("IDShipping").Value)%>"><i class="fa fa-trash-alt"></i></a>
      <a href="manage_shippingmethods.asp?Edit=yes&ID=<%=(rsGetShipping.Fields.Item("IDShipping").Value)%>" target="_top" class="LeftNavLinks"><%=(rsGetShipping.Fields.Item("ShippingName").Value)%></a></td>
      <td><%=(rsGetShipping.Fields.Item("ShippingAmount").Value)%></td>
      <td><p><%=(rsGetShipping.Fields.Item("ShippingDesc_Public").Value)%></p>
      <p><strong>Notes:</strong> <%=(rsGetShipping.Fields.Item("ShippingDesc_Private").Value)%></p></td>
      <td><strong><%=(rsGetShipping.Fields.Item("ShippingWeight").Value)%> oz restriction</strong><br>
      Show on FAQ page:&nbsp;
      <% if (rsGetShipping.Fields.Item("FAQShow").Value) = 0 then %>No<% else %>Yes<% end if %><br>
      Sorting order: <%=(rsGetShipping.Fields.Item("SortOrder").Value)%><br>
      Subtotal discount: $<%=(rsGetShipping.Fields.Item("DiscountSubtotal").Value)%> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
      Discount amount: $<%=(rsGetShipping.Fields.Item("ShippingDiscount").Value)%> </td>
    </tr>
    <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsGetShipping.MoveNext()
Wend
%>
</table>
</div>
</body>
</html>

<%
rsGetShipping.Close()
Set rsGetShipping = Nothing
%>
<%
rsGetONE.Close()
Set rsGetONE = Nothing
%>

