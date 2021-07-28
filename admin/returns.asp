<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
'==== PAGE HAS BEEN BOOSTRAPPED =======
bootstrapped = "yes"


If request.querystring("status") <> "" then
	var_status =  request.querystring("status")
else
	var_status = ""
end if

If request.querystring("sortby") <> "" then
	sortby =  request.querystring("sortby")
else
	sortby = "date_order_placed"
end if

set rsGetRecords = Server.CreateObject("ADODB.Recordset")
rsGetRecords.ActiveConnection = MM_bodyartforms_sql_STRING
rsGetRecords.Source = "SELECT ID, customer_first, customer_last, our_notes, Comments_OrderError, date_sent, pay_method, email, country FROM sent_items WHERE shipped = '" + var_status + "' ORDER BY " + sortby + " ASC"
rsGetRecords.CursorLocation = 3 'adUseClient
rsGetRecords.LockType = 1 'Read-only records
rsGetRecords.Open()
%>
<html>
<head>
<title>Orders by status</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"></head>
<body>

  <!--#include file="admin_header.asp"-->
<div class="p-3">
<h4>
	Order set to <%= request.querystring("status") %>
</h4>

  <form name="form1" method="get" action="returns.asp">
      <input type="radio" name="status" id="status" value="ON HOLD">
      <label class="mr-4" for="status">On hold</label>
    
    <input type="radio" name="status" id="status" value="RETURN">
    <label class="mr-4">Return</label>
      <input type="radio" name="status" id="status" value="Waiting for PayPal eCheck to clear">
      <label class="mr-4">Paypal eCheck</label>
      <input type="radio" name="status" id="status" value="PACKAGE CAME BACK">
      <label class="mr-4">Package came back</label>
      <input type="radio" name="status" id="status" value="RETURN (EXCEPTION)">
      <label class="mr-4">Return (Exception)</label>
       <input type="radio" name="status" id="status" value="Flagged">
       <label class="mr-4">Flagged</label>
       <input type="radio" name="status" id="status" value="Chargeback">
       <label class="mr-4">Chargeback</label>

    
  
    <div class="mt-2">
      <button class="btn btn-sm btn-secondary mr-5" type="submit" name="button" id="button">Submit</button>
      <a class="mr-3" href="returns.asp?status=<%= request.querystring("status") %>&sortby=date_sent" class="HomePageLinks">Sort 
     by date (default)</a>
     <a href="returns.asp?status=<%= request.querystring("status") %>&sortby=customer_first" class="HomePageLinks">Sort by 
    name A-Z</a>
    </div>
</form>
  
  <table class="table table-sm table-striped table-hover">
<% While NOT rsGetRecords.EOF %>
  <tr align="left" valign="middle" > 
    <td width="8%" valign="top"><p><a href="invoice.asp?ID=<%=(rsGetRecords.Fields.Item("ID").Value)%>" target="_blank"><strong><%=(rsGetRecords.Fields.Item("ID").Value)%></strong></a><br>
    <%=(rsGetRecords.Fields.Item("customer_first").Value)%>&nbsp;<%=(rsGetRecords.Fields.Item("customer_last").Value)%></p>
<p><a href="order history.asp?var_first=<%=(rsGetRecords.Fields.Item("customer_first").Value)%>&var_last=<%=(rsGetRecords.Fields.Item("customer_last").Value)%>" target="_blank">View history</a> </p>
      <p><%=(rsGetRecords.Fields.Item("date_sent").Value)%><br>
        <%=(rsGetRecords.Fields.Item("pay_method").Value)%><br>
      <%=(rsGetRecords.Fields.Item("country").Value)%></p></td>
    <td width="52%" valign="top">
      <% if rsGetRecords.Fields.Item("our_notes").Value <> "" then %>
        <%=Replace(rsGetRecords.Fields.Item("our_notes").Value, vbCrLF, "<br />" + vbCrLF)%>
      <% end if %>
      <br><%=(rsGetRecords.Fields.Item("Comments_OrderError").Value)%> <br>
      <%
Dim rsGetOrderDetails
Dim rsGetOrderDetails_numRows

Set rsGetOrderDetails = Server.CreateObject("ADODB.Recordset")
With rsGetOrderDetails
rsGetOrderDetails.ActiveConnection = MM_bodyartforms_sql_STRING
rsGetOrderDetails.Source = "SELECT OrderDetailID, qty, title, ProductDetail1, item_price, notes  FROM dbo.QRY_OrderDetails  WHERE ID = " & rsGetRecords.Fields.Item("ID").Value & ""
rsGetOrderDetails.CursorLocation = 3 'adUseClient
rsGetOrderDetails.LockType = 1 'Read-only records
rsGetOrderDetails.Open()

LineItem = 0
SumLineItem = 0

Do While Not.Eof
%>
      <%=(rsGetOrderDetails.Fields.Item("qty").Value)%> | <%=(rsGetOrderDetails.Fields.Item("title").Value)%>&nbsp; <%=(rsGetOrderDetails.Fields.Item("ProductDetail1").Value)%> &nbsp;<b><%=(rsGetOrderDetails.Fields.Item("notes").Value)%></b><br>
      <%
LineItem = rsGetOrderDetails.Fields.Item("item_price").Value * rsGetOrderDetails.Fields.Item("qty").Value
SumLineItem = SumLineItem + LineItem

.Movenext()
Loop
End With 

rsGetOrderDetails.Close()
Set rsGetOrderDetails = Nothing
%>    </td>
  </tr>
  <% 
  rsGetRecords.MoveNext()
Wend
%>
</table>
</div>
</body>
</html>
<%
rsGetRecords.Close()
%>
