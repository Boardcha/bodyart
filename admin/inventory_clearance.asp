<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
'==== PAGE HAS BEEN BOOSTRAPPED =======
bootstrapped = "yes"


set rsCheck = Server.CreateObject("ADODB.Recordset")
rsCheck.ActiveConnection = MM_bodyartforms_sql_STRING
rsCheck.Source = "SELECT * FROM dbo.inventory  WHERE (type = 'Clearance' OR type = 'onetime' OR type = 'blowout' OR type LIKE '%limited%' OR type = 'Discontinued' OR type = 'One time buy' OR type = 'Consignment') AND qty <= 0  AND item_active = 1 AND product_active = 1 ORDER BY title ASC"
rsCheck.CursorLocation = 3 'adUseClient
rsCheck.LockType = 1 'Read-only records
rsCheck.Open()
rsCheck_numRows = 0

Dim Repeat1__numRows
Repeat1__numRows = -1
Dim Repeat1__index
Repeat1__index = 0
rsCheck_numRows = rsCheck_numRows + Repeat1__numRows

' Set all clearance and limited items to inactive if they are out
if request.querystring("UpdateAll") = "yes" then

Do While ((Repeat1__numRows <> 0) AND (NOT rsCheck.EOF))

	set rsItemsInactive = Server.CreateObject("ADODB.Command")
	rsItemsInactive.ActiveConnection = MM_bodyartforms_sql_STRING
	rsItemsInactive.CommandText = "UPDATE inventory SET item_active = 0 WHERE (type = 'Clearance' OR type = 'onetime' OR type LIKE '%limited%' OR type = 'blowout' OR type = 'One time buy' OR type = 'Discontinued' OR type = 'Consignment') AND qty <= 0  AND item_active = 1 AND product_active = 1" 
	rsItemsInactive.Execute()
	
Repeat1__index=Repeat1__index+1
Repeat1__numRows=Repeat1__numRows-1
rsCheck.MoveNext()
Loop

	set SetProduct = Server.CreateObject("ADODB.Command")
	SetProduct.ActiveConnection = MM_bodyartforms_sql_STRING
	SetProduct.CommandText = "UPDATE jewelry SET jewelry.active = 0 WHERE jewelry.active = 1 AND (jewelry.type = 'Clearance' OR jewelry.type = 'onetime' OR jewelry.type LIKE '%limited%' OR jewelry.type = 'blowout' OR jewelry.type = 'One time buy' OR jewelry.type = 'Discontinued' OR jewelry.type = 'Consignment') AND NOT EXISTS (SELECT 1 FROM ProductDetails WHERE ProductDetails.ProductID = jewelry.ProductID AND ProductDetails.active = 1)"
	SetProduct.Execute()

end if 
%>
<html>
<head>
<title>Inventory check</title>
</head>
<body>
<!--#include file="admin_header.asp"-->

<div class="p-3">

  <table class="table table-sm table-hover table-striped" id="details-table">
    <thead class="thead-dark">
	<tr>
      <th colspan="8"><form name="form1" method="post" action="inventory_clearance.asp?UpdateAll=yes">
        Set all items to inactive &nbsp;&nbsp;&nbsp;
            <button class="btn btn-sm btn-secondary" type="submit">Go <i class="fa fa-lg fa-angle-double-right"></i></button>
            </form>
      </th>
    </tr>
    <tr>
      <th>Product</th>
	  <th>Brand</th>
	  <th>Code</th>
	  <th>Bin</th>
	  <th>Section</th>
	  <th>Location</th>
	  <th>Discount</th>
	  <th>Added</th>
    </tr>
	</thead>
        <% 
rsCheck.ReQuery
	Repeat1__numRows = -1
Repeat1__index = 0
While ((Repeat1__numRows <> 0) AND (NOT rsCheck.EOF)) 
%>
    <tr>
           <td>
			<a href="product-edit.asp?ProductID=<%=(rsCheck.Fields.Item("ProductID").Value)%>&info=less">
				<img class="mr-2" style="width:90px;height:90px" src="https://bodyartforms-products.bodyartforms.com/<%= rsCheck("picture") %>" alt="Product photo">
				<%=(rsCheck.Fields.Item("title").Value)%> - <%=(rsCheck.Fields.Item("Gauge").Value)%> <%=(rsCheck.Fields.Item("Length").Value)%> <%=(rsCheck.Fields.Item("ProductDetail1").Value)%></a></td>
		   <td><%= rsCheck("brandname") %></td>
			<td>
				<%=(rsCheck.Fields.Item("detail_code").Value)%>
			</td>
			<td>
				<%=(rsCheck.Fields.Item("BinNumber_Detail").Value)%>
			</td>
			<td>
				<%=(rsCheck.Fields.Item("ID_Description").Value)%>
			</td>
			<td>
				<%=(rsCheck.Fields.Item("location").Value)%>
			</td>
			<td><%= rsCheck("SaleDiscount") %>%</td>
			<td><%= FormatDateTime(rsCheck("date_added"),vbShortDate) %></td>
    </tr>
    <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsCheck.MoveNext()
Wend
%>
</table>
</div>
</body>
</html>
<%
rsCheck.Close()
Set rsCheck = Nothing
%>