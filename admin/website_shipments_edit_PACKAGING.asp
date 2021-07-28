<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
Dim MM_editAction
MM_editAction = CStr(Request.ServerVariables("SCRIPT_NAME"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Server.HTMLEncode(Request.QueryString)
End If

' boolean to abort record edit
Dim MM_abortEdit
MM_abortEdit = false
%>
<%
' IIf implementation
Function MM_IIf(condition, ifTrue, ifFalse)
  If condition = "" Then
    MM_IIf = ifFalse
  Else
    MM_IIf = ifTrue
  End If
End Function
%>
<%
If (CStr(Request("MM_update")) = "FRM_update") Then
  If (Not MM_abortEdit) Then
    ' execute the update
    Dim MM_editCmd

    Set MM_editCmd = Server.CreateObject ("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_bodyartforms_sql_STRING
    MM_editCmd.CommandText = "UPDATE dbo.sent_items SET our_notes = ?, coupon_amt = ?, item_description = ? WHERE ID = ?" 
    MM_editCmd.Prepared = true
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param1", 203, 1, 1073741823, Request.Form("our_notes")) ' adLongVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param2", 5, 1, -1, MM_IIF(Request.Form("coupon_amt"), Request.Form("coupon_amt"), null)) ' adDouble
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param3", 203, 1, 1073741823, Request.Form("order")) ' adLongVarWChar
    MM_editCmd.Parameters.Append MM_editCmd.CreateParameter("param4", 5, 1, -1, MM_IIF(Request.Form("MM_recordId"), Request.Form("MM_recordId"), null)) ' adDouble
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    ' append the query string to the redirect URL
    Dim MM_editRedirectUrl
    MM_editRedirectUrl = "website_shipments_edit_PACKAGING.asp"
    If (Request.QueryString <> "") Then
      If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0) Then
        MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
      Else
        MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
      End If
    End If
Response.Redirect "website_shipments_edit_PACKAGING.asp?ID="+ Request.Form("MM_recordId")+""
  End If
End If
%>
<%
set rsAddEbay = Server.CreateObject("ADODB.Recordset")
rsAddEbay.ActiveConnection = MM_bodyartforms_sql_STRING
rsAddEbay.Source = "SELECT *  FROM sent_items  WHERE (ID = '" + Request.Form("invoice_num") + "') OR (ID = '" + Request.QueryString("ID") + "') OR (ID = '" + Request.Form("MM_recordId") + "') OR (transactionID = '" + Request.Form("TransID") + "')"
rsAddEbay.CursorLocation = 3 'adUseClient
rsAddEbay.LockType = 1 'Read-only records
rsAddEbay.Open()
rsAddEbay_numRows = 0
%>
<%
Dim rsGetOrderItems__MMColParam
rsGetOrderItems__MMColParam = "1"
If (Request.QueryString("ID") <> "") Then 
  rsGetOrderItems__MMColParam = Request.QueryString("ID")
End If
%>
<%
Dim rsGetOrderItems
Dim rsGetOrderItems_numRows

Set rsGetOrderItems = Server.CreateObject("ADODB.Recordset")
rsGetOrderItems.ActiveConnection = MM_bodyartforms_sql_STRING
rsGetOrderItems.Source = "SELECT * FROM dbo.QRY_OrderDetails WHERE InvoiceID = '" & rsAddEbay.Fields.Item("ID").Value & "' ORDER BY OrderDetailID ASC"
' changed by MaximumASP to cursor type 1 (keyset) and a lock type of 3 (optimistic)
' originally set to 0 (forward-only) and a lock type of 1 (read only)
rsGetOrderItems.CursorLocation = 3 'adUseClient
rsGetOrderItems.LockType = 1 'Read-only records
rsGetOrderItems.Open()

rsGetOrderItems_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
rsGetOrderItems_numRows = rsGetOrderItems_numRows + Repeat1__numRows
%>
<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

Dim rsGetOrderItems_total
Dim rsGetOrderItems_first
Dim rsGetOrderItems_last

' set the record count
rsGetOrderItems_total = rsGetOrderItems.RecordCount

' set the number of rows displayed on this page
If (rsGetOrderItems_numRows < 0) Then
  rsGetOrderItems_numRows = rsGetOrderItems_total
Elseif (rsGetOrderItems_numRows = 0) Then
  rsGetOrderItems_numRows = 1
End If

' set the first and last displayed record
rsGetOrderItems_first = 1
rsGetOrderItems_last  = rsGetOrderItems_first + rsGetOrderItems_numRows - 1

' if we have the correct record count, check the other stats
If (rsGetOrderItems_total <> -1) Then
  If (rsGetOrderItems_first > rsGetOrderItems_total) Then
    rsGetOrderItems_first = rsGetOrderItems_total
  End If
  If (rsGetOrderItems_last > rsGetOrderItems_total) Then
    rsGetOrderItems_last = rsGetOrderItems_total
  End If
  If (rsGetOrderItems_numRows > rsGetOrderItems_total) Then
    rsGetOrderItems_numRows = rsGetOrderItems_total
  End If
End If
%>
<%
' *** Recordset Stats: if we don't know the record count, manually count them

If (rsGetOrderItems_total = -1) Then

  ' count the total records by iterating through the recordset
  rsGetOrderItems_total=0
  While (Not rsGetOrderItems.EOF)
    rsGetOrderItems_total = rsGetOrderItems_total + 1
    rsGetOrderItems.MoveNext
  Wend

  ' reset the cursor to the beginning
  If (rsGetOrderItems.CursorType > 0) Then
    rsGetOrderItems.MoveFirst
  Else
    rsGetOrderItems.Requery
  End If

  ' set the number of rows displayed on this page
  If (rsGetOrderItems_numRows < 0 Or rsGetOrderItems_numRows > rsGetOrderItems_total) Then
    rsGetOrderItems_numRows = rsGetOrderItems_total
  End If

  ' set the first and last displayed record
  rsGetOrderItems_first = 1
  rsGetOrderItems_last = rsGetOrderItems_first + rsGetOrderItems_numRows - 1
  
  If (rsGetOrderItems_first > rsGetOrderItems_total) Then
    rsGetOrderItems_first = rsGetOrderItems_total
  End If
  If (rsGetOrderItems_last > rsGetOrderItems_total) Then
    rsGetOrderItems_last = rsGetOrderItems_total
  End If

End If
%>
<%
' *** FX Update Multiple Records in FRM_OrderItems
If (rsGetOrderItems_first <> "") Then upd_rsGetOrderItems = rsGetOrderItems_first-1 Else upd_rsGetOrderItems = 0 End If ' counter
If (cStr(Request.Form("Submit5")) <> "") Then
  FX_sqlerror = ""
  FX_updredir = "website_shipments_edit_PACKAGING.asp?ID="+request.form("invoice_num")+""
  tmp = "ADODB.Command"
  Set update_Multi = Server.CreateObject(tmp)
  update_Multi.ActiveConnection = MM_bodyartforms_sql_STRING
  For N = upd_rsGetOrderItems+1 To rsGetOrderItems_total
      If (Request.Form("notes"&N) <> "") Then s7 = Replace(Request.Form("notes"&N),"'","''") Else s7 = "" End If
      If (Request.Form("fx_updmatch"&N) <> "") Then sw = Replace(Request.Form("fx_updmatch"&N),"'","''") Else sw = "0" End If
    On Error Resume Next
      update_Multi.CommandText = "UPDATE dbo.TBL_OrderSummary SET notes='"+s7+"' WHERE OrderDetailID="+sw+""
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
    FX_updredir = FX_updredir
  End If
  If (FX_sqlerror <> "") Then
    Response.Write("<font color=""red"">"&FX_sqlerror&"</font>")
  Else Response.Redirect(FX_updredir) End If
End If
%>
<html>
<head>
<link href="../CSS/Admin.css" rel="stylesheet" type="text/css" />
<title>Edit order</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body>
<!--#include file="admin_header.asp"-->
<br>
<form ACTION="<%=MM_editAction%>" METHOD="POST" name="FRM_update" class="form">
  <div class="ProductsHeader">
  <div style="float: left;">
    INVOICE # <%=(rsAddEbay.Fields.Item("ID").Value)%> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span class="PagingLinks_Inactive">Placed on <%=(rsAddEbay.Fields.Item("date_order_placed").Value)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Packaged by: <% = (rsAddEbay.Fields.Item("PackagedBy").Value) %><% if (rsAddEbay.Fields.Item("ScanInvoice_Timestamp").Value) <> "" then %>&nbsp;&nbsp;Has been  packaged
  <% end if %></span></div>
  <div style="clear: both;"></div>
  </div><!--End product header-->
  <div class="ContentText">
  <div style="float: left; width: 48%;">

    <p><a href="invoices/print-friendly-invoice.asp?ID=<%=(rsAddEbay.Fields.Item("ID").Value)%>" target="_blank">Print invoice</a></p>
    <p>PRIVATE NOTES:<br>
      <textarea name="our_notes" cols="60" rows="5" class="RetailInput" id="our_notes"><%=(rsGetUser.Fields.Item("name").Value)%>&nbsp;<%= now() %>&#x000D;&#x000D;<%=(rsAddEbay.Fields.Item("our_notes").Value)%></textarea>
      <br>
    </p>
    <p><strong><%=(rsAddEbay.Fields.Item("customer_comments").Value)%></strong></p>
    <p>&nbsp;</p>
    <p><span style="clear: both;">
      <input name="Submit" type="submit" value="UPDATE MAIN ORDER">
      <input name="coupon_amt" type="hidden" value="0">
      <input type="hidden" name="MM_update" value="FRM_update">
      <input type="hidden" name="MM_recordId" value="<%= rsAddEbay.Fields.Item("ID").Value %>">
    </span></p>
  </div>
  <!--  End left float -->
  <div style="float: right; width: 48%;">
  <div class="AccountPageHeaders">Shipping address</div>
  <div class="AccountPageContent">
    <% if (rsAddEbay.Fields.Item("company").Value) <> "" then %>
    <%=(rsAddEbay.Fields.Item("company").Value)%><br>
    <% end if %>
    <%=(rsAddEbay.Fields.Item("customer_first").Value)%> &nbsp;<%=(rsAddEbay.Fields.Item("customer_last").Value)%><br>
    <%=(rsAddEbay.Fields.Item("address").Value)%> <br>
    <% if (rsAddEbay.Fields.Item("address2").Value) <> "" then %>
    <%=(rsAddEbay.Fields.Item("address2").Value)%> <br>
    <% end if %>
    <%=(rsAddEbay.Fields.Item("city").Value)%>, <%=(rsAddEbay.Fields.Item("state").Value)%><%=(rsAddEbay.Fields.Item("province").Value)%>&nbsp;&nbsp;<%=(rsAddEbay.Fields.Item("zip").Value)%><br>
    <%=(rsAddEbay.Fields.Item("country").Value)%>
  </div>
  <fieldset>
    <br>
    
  </fieldset>
    
    
    PUBLIC notes and info to print on invoice: <br>
    <textarea name="order" cols="60" rows="3" class="form_fieldpadding"><%=(rsAddEbay.Fields.Item("item_description").Value)%>
      </textarea>
    
  </div>
  <!--  End right float -->
  <div style="clear: both;" align="center"></div>
  </div><!--ENd content div-->
  </div>
</form>

<br>
<form action="<%=MM_editAction%>" method="POST" name="FRM_OrderItems">
  
  <div class="ProductsHeader">
  <div style="float: left;">
  <%=(rsGetOrderItems_total)%> items in order</div>
  <div style="clear: both;"></div>
  </div>
<div class="ContentText">
  <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#DDDDDD" class="ActiveTable">
    <% 
LineItem = 0
SumLineItem = 0

While ((Repeat1__numRows <> 0) AND (NOT rsGetOrderItems.EOF)) 
%>
    <% upd_rsGetOrderItems = upd_rsGetOrderItems+1 %>
    <%If (Repeat1__numRows Mod 2) Then%>
    <tr valign="middle" bgcolor="#ececec">
      <%End If%>
 
        <td width="20%" valign="top">&nbsp;</td>
      <td width="30%" valign="top"><strong><%=(rsGetOrderItems.Fields.Item("ProductDetailID").Value)%></strong>&nbsp;&nbsp;<a href="product-edit.asp?ProductID=<%=(rsGetOrderItems.Fields.Item("ProductID").Value)%>&info=less" class="HomePageLinks"><%=(rsGetOrderItems.Fields.Item("title").Value)%></a>&nbsp;&nbsp;<%=(rsGetOrderItems.Fields.Item("ProductDetail1").Value)%>&nbsp;<%=(rsGetOrderItems.Fields.Item("Gauge").Value)%>&nbsp;<%=(rsGetOrderItems.Fields.Item("Length").Value)%></td>
      <td align="right" valign="top"><input type="hidden" name="fx_updmatch<%=upd_rsGetOrderItems%>" value="<%=(rsGetOrderItems.Fields.Item("OrderDetailID").Value)%>">
        <input name="notes<%=upd_rsGetOrderItems%>" type="text" class="adminfields" id="notes" size="30" value="<% if isNull(rsGetOrderItems.Fields.Item("notes").Value) then %><% else %><%= Server.HTMLEncode(rsGetOrderItems.Fields.Item("notes").Value) %><% end if %>"></td>
      </tr>
    <% 
LineItem = rsGetOrderItems.Fields.Item("item_price").Value * rsGetOrderItems.Fields.Item("qty").Value
SumLineItem = SumLineItem + LineItem

  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsGetOrderItems.MoveNext()
Wend
%>
    </table>
<input type="submit" name="Submit5" value="Update order items">
  <input name="invoice_num" type="hidden" value="<%=(rsAddEbay.Fields.Item("ID").Value)%>">  </div><!--ENd content div-->
</form>

<br>
</body>
</html>
<%
rsAddEbay.Close()
%>
<%
rsGetOrderItems.Close()
Set rsGetOrderItems = Nothing
%>
<%
rsGetUser.Close()
Set rsGetUser = Nothing
%>
