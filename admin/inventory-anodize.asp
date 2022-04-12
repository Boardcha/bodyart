<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
'==== PAGE HAS BEEN BOOSTRAPPED =======
bootstrapped = "yes"

if request.querystring("sortby") <> "" then 
    session("anodize_sort") = request.querystring("sortby")
end if

if session("anodize_sort") <> "" then
    sortby = session("anodize_sort")
end if

if sortby = "needed" then
	var_sortby = "needed DESC"
elseif sortby = "color" then
	var_sortby = "ProductDetail1 ASC"	
else
	var_sortby = "ProductDetail1 ASC"
end if

set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn  
objCmd.CommandText = "SELECT ProductDetails.ProductDetailID, jewelry.title, ProductDetails.ProductDetail1, ProductDetails.qty, stock_qty, restock_threshold,  CASE WHEN qty < restock_threshold THEN  stock_qty - qty ELSE 0 END as 'needed', jewelry.ProductID, ProductDetails.Gauge, ProductDetails.Length, jewelry.picture, ProductDetails.location,  TBL_Barcodes_SortOrder.ID_Description, ProductDetails.BinNumber_Detail, colors FROM ProductDetails INNER JOIN jewelry ON ProductDetails.ProductID = jewelry.ProductID INNER JOIN TBL_Barcodes_SortOrder ON ProductDetails.DetailCode = TBL_Barcodes_SortOrder.ID_Number  WHERE jewelry.tags LIKE '%bulkanodize%' ORDER BY CASE WHEN type = 'None' THEN 0 ELSE 1 END ASC, " & var_sortby

set rsGetAnodizedList = Server.CreateObject("ADODB.Recordset")
rsGetAnodizedList.CursorLocation = 3 'adUseClient
rsGetAnodizedList.Open objCmd
rsGetAnodizedList.PageSize = 100
total_records = rsGetAnodizedList.RecordCount
intPageCount = rsGetAnodizedList.PageCount

' Variables for paging
Select Case Request("Action")
    case "<<"
        intpage = 1
    case "<"
        intpage = Request("intpage")-1
        if intpage < 1 then intpage = 1
    case ">"
        intpage = Request("intpage")+1
        if intpage > intPageCount then intpage = IntPageCount
    Case ">>"
        intpage = intPageCount
    case else
        intpage = 1
end select	
%>
<html>
<head>
<title>Anodized Products List</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body>
<!--#include file="admin_header.asp"-->
<div class="p-3">
<h5>
    Anodized Products List
</h5>
<a href="?sortby=needed">Sort by amount needed</a> | <a href="?sortby=color">Sort by color (default)</a>
<div class="container-fluid p-0 m-0">
	<div class="row">
		<div class="col">
			<!--<form class="form-inline" action="?po_id=<%= request("po_id") %>" method="post">
				<select class="form-control form-control-sm mr-3" name="filter" id="filter">
				<option value="0" selected>View all (default)</option>
   
				</select>
				<input class="btn btn-sm btn-secondary" type="submit" value="Filter">
			</form>-->
		</div>
		<div class="col-auto">
		</div>
	</div><!-- row -->
</div><!-- container -->

<% If Not rsGetAnodizedList.EOF Or Not rsGetAnodizedList.BOF Then %>
<!--#include file="includes/inc-paging.asp"-->
<table  class="table table-sm table-striped table-hover mt-2">
	<thead class="thead-dark">  
	<tr>
            <th class="sticky-top">Detail #</th>
            <th class="sticky-top">Location</th>
            <th class="sticky-top">In stock</th>
            <th class="sticky-top">Max qty</th>
            <th class="sticky-top">Low threshold</th>
            <th class="sticky-top">Amt needed</th>
			<th class="sticky-top">Description</th>
			<th class="sticky-top">Gauge</th>
            <th class="sticky-top">Size</th>
            <th class="sticky-top">Color</th>
		  </tr>
		</thead>	
<% 
rsGetAnodizedList.AbsolutePage = intPage '======== PAGING
For intRecord = 1 To rsGetAnodizedList.PageSize 


var_qty_warning = ""
if rsGetAnodizedList.Fields.Item("qty").Value < rsGetAnodizedList.Fields.Item("restock_threshold").Value then
        var_qty_warning = "bg-warning py-0 px-2 font-weight-bold m-0 rounded"
end if

 %>
 <tr>
     <td>
    <a href="product-edit.asp?ProductID=<%=(rsGetAnodizedList.Fields.Item("ProductID").Value)%>" target="_blank">
        <%=(rsGetAnodizedList.Fields.Item("ProductDetailID").Value)%></a> 
    </td>
<td>
    <%=(rsGetAnodizedList.Fields.Item("location").Value)%> - <%= replace(rsGetAnodizedList.Fields.Item("ID_Description").Value, "Main", "")%>
<% if rsGetAnodizedList.Fields.Item("BinNumber_Detail").Value <> 0 then %>
			 BIN <%=(rsGetAnodizedList.Fields.Item("BinNumber_Detail").Value)%>
			<% end if %>
</td>
<td>
    <span class="<%= var_qty_warning %>"><%=(rsGetAnodizedList.Fields.Item("qty").Value)%></span>
</td>
<td>
    <%= rsGetAnodizedList.Fields.Item("stock_qty").Value %>
</td>
<td>
    <%= rsGetAnodizedList("restock_threshold") %>
</td>
<td>
    <% if rsGetAnodizedList.Fields.Item("needed").Value > 0 then %>
    <span class='alert alert-info py-0 px-2 font-weight-bold m-0'><%= rsGetAnodizedList.Fields.Item("needed").Value %></span>
    <% end if %>
</td>
<td>
    <%=(rsGetAnodizedList.Fields.Item("title").Value)%>
</td>
<td >
	<%=(rsGetAnodizedList.Fields.Item("gauge").Value)%>
</td>
<td >
	<%=(rsGetAnodizedList.Fields.Item("length").Value)%>
</td>
<td>
    <%=(rsGetAnodizedList.Fields.Item("ProductDetail1").Value)%>     

  </td>
  </tr>               <% 
  rsGetAnodizedList.MoveNext()
  If rsGetAnodizedList.EOF Then Exit For  ' ====== PAGING
  Next ' ====== PAGING
%>
 </table>
<!--#include file="includes/inc-paging.asp"-->

          <% End If ' end Not rsGetAnodizedList.EOF Or NOT rsGetAnodizedList.BOF %>
    

</div><!--admin content-->
</body>
</html>
<%
DataConn.Close()
Set DataConn = Nothing
%>
