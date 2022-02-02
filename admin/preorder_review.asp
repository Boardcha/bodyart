<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
'==== PAGE HAS BEEN BOOSTRAPPED =======
bootstrapped = "yes"


Dim rsGetPreorders
Dim rsGetPreorders_numRows

Set rsGetPreorders = Server.CreateObject("ADODB.Recordset")
rsGetPreorders.ActiveConnection = MM_bodyartforms_sql_STRING
rsGetPreorders.Source = "SELECT TOP (100) PERCENT dbo.TBL_OrderSummary.InvoiceID, dbo.TBL_OrderSummary.ProductID, dbo.TBL_OrderSummary.DetailID, dbo.TBL_OrderSummary.qty, dbo.TBL_OrderSummary.PreOrder_Desc, dbo.jewelry.title, dbo.ProductDetails.ProductDetail1, dbo.sent_items.ID, dbo.TBL_OrderSummary.OrderDetailID, dbo.jewelry.customorder, dbo.sent_items.shipped, dbo.jewelry.brandname, dbo.TBL_OrderSummary.item_shipped, dbo.TBL_OrderSummary.item_ordered,  dbo.TBL_OrderSummary.item_received, dbo.TBL_OrderSummary.notes, total_items_subtotal, dbo.TBL_OrderSummary.status, dbo.ProductDetails.ProductDetailID,  dbo.sent_items.customer_first, dbo.ProductDetails.Gauge, dbo.ProductDetails.Length, dbo.sent_items.customer_comments FROM dbo.TBL_OrderSummary INNER JOIN  dbo.jewelry ON dbo.TBL_OrderSummary.ProductID = dbo.jewelry.ProductID INNER JOIN  dbo.ProductDetails ON dbo.TBL_OrderSummary.DetailID = dbo.ProductDetails.ProductDetailID INNER JOIN  dbo.sent_items ON dbo.TBL_OrderSummary.InvoiceID = dbo.sent_items.ID WHERE (dbo.jewelry.customorder = 'yes') AND (dbo.TBL_OrderSummary.item_ordered = 0) AND (dbo.sent_items.shipped = N'CUSTOM ORDER IN REVIEW') ORDER BY dbo.TBL_OrderSummary.InvoiceID"
rsGetPreorders.CursorLocation = 3 'adUseClient
rsGetPreorders.LockType = 1 'Read-only records
rsGetPreorders.Open()


if request.form("FRMupdate") = "yes" then
temp = Replace( Request.Form("Checkbox"), "'", "''" ) 
varID = Split( temp, ", " ) 

set commUpdate = Server.CreateObject("ADODB.Command")
commUpdate.ActiveConnection = MM_bodyartforms_sql_STRING

For i = 0 To UBound(varID) 

commUpdate.CommandText = "UPDATE dbo.QRY_OrderDetails SET shipped = 'CUSTOM ORDER APPROVED' WHERE OrderDetailID = " & varID(i) 
    ' comment out next line AFTER IT WORKS 
    'Response.Write "DEBUG SQL: " & commUpdate.CommandText & "<BR/>" 
commUpdate.Execute()

   Next

   Response.Redirect("preorder_review.asp")
end if
%>
<html>
<head>
<title>Custom order review &amp; approval</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"></head>
<body>
<!--#include file="admin_header.asp"-->
<div class="p-3">

<h5>Review &amp; approve custom orders </h5> 

<form action="" method="post" name="FRM_update" id="FRM_update">
<% If Not rsGetPreorders.EOF Then %>
<table class="table table-sm table-striped table-hover">
	<thead class="thead-dark">
	<tr> 
		<th width="5%"> </th>
		<th width="30%">Invoice</th>
		<th width="5%">Qty</th>
		<th width="60%">Description</th>
	</tr>
	</thead>
              <% 
While NOT rsGetPreorders.EOF 
%>
                <tr>
					<td class="text-center align-middle"><input name="Checkbox" type="checkbox" value="<%=(rsGetPreorders.Fields.Item("OrderDetailID").Value)%>">
                    </td>
                  <td>
                    <a class="mr-2" href="invoice.asp?ID=<%=(rsGetPreorders.Fields.Item("InvoiceID").Value)%>"><%=(rsGetPreorders.Fields.Item("InvoiceID").Value)%></a> 
                    <% if rsGetPreorders("total_items_subtotal") = 0 then %>
                    <span class="badge badge-warning">Fetching total tomorrow</span>
                    <% elseif rsGetPreorders("total_items_subtotal") > 150 then %>
                    <span class="badge badge-danger">Over $150</span>
                    <% end if %>
                    <a class="d-block mt-2" href="email_template_send.asp?ID=<%=(rsGetPreorders.Fields.Item("InvoiceID").Value)%>&type=generic">Email <%=(rsGetPreorders.Fields.Item("customer_first").Value)%></a></td>
                  <td class="align-middle"><%=(rsGetPreorders.Fields.Item("qty").Value)%></td>
                  <td><b><%=(rsGetPreorders.Fields.Item("brandname").Value)%></b>&nbsp;&nbsp;&nbsp;<a href="product-edit.asp?ProductID=<%=(rsGetPreorders.Fields.Item("ProductID").Value)%>&info=less"><%=(rsGetPreorders.Fields.Item("title").Value)%>&nbsp;<%=(rsGetPreorders.Fields.Item("gauge").Value)%>&nbsp;<%=(rsGetPreorders.Fields.Item("Length").Value)%>&nbsp;<%=(rsGetPreorders.Fields.Item("ProductDetail1").Value)%></a><br>
                  Specs: <% if (rsGetPreorders.Fields.Item("PreOrder_Desc").Value) <> "" then %><%=Server.HTMLEncode(rsGetPreorders.Fields.Item("PreOrder_Desc").Value)%><% end if %>
                <% if rsGetPreorders("customer_comments") <> "" then %>
                    <div class="badge badge-info"><%= rsGetPreorders("customer_comments") %></div>
                <% end if %>
                
                </td>
                </tr>
                <% 
  rsGetPreorders.MoveNext()
Wend
%>

          </table>

          <div class="text-center">
            <input type="submit" name="Submit2" value="Approve orders" class="btn btn-primary">
            <input name="FRMupdate" type="hidden" id="FRMupdate" value="yes">
          </div>
        
        <% End If ' end Not rsGetPreorders.EOF Or NOT rsGetPreorders.BOF %>
      <% If rsGetPreorders.EOF And rsGetPreorders.BOF Then %>
        <p>No custom orders to review </p>
        <% End If ' end rsGetPreorders.EOF And rsGetPreorders.BOF %>
</td>
  </tr>
</table>
</form>
</div>
</body>
</html>
<%
rsGetPreorders.Close()
Set rsGetPreorders = Nothing
%>
