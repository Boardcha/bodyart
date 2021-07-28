<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
'==== PAGE HAS BEEN BOOSTRAPPED =======
bootstrapped = "yes"


set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT * FROM jewelry WHERE to_be_pulled = 1 AND pull_completed = 1"
set rsGetProducts = Server.CreateObject("ADODB.Recordset")
rsGetProducts.CursorLocation = 3 'adUseClient
rsGetProducts.Open objCmd
rsGetProducts.PageSize = 1
total_records = rsGetProducts.RecordCount
intPageCount = rsGetProducts.PageCount

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
<title>Review pulled discontinued items</title>
</head>
<body>

<!--#include file="admin_header.asp"-->
<div class="m-3">
<% If Session("SubAccess") <> "N" then ' DISPLAY ONLY TO PEOPLE WHO HAVE ACCESS TO THIS PAGE %>

<h5><%= total_records %> products to review</h5>
<nav>
	<ul class="pagination" style="justify-content: center">
<%	if Intpage <= intpagecount then
	if intpagecount <> 1 then 
	 if Intpage <> 1 then %>
<li class="page-item"><a class="page-link text-ltpurple"  href="?action=<<&amp;intpage=<%=intpage%>&amp;userID=<%= request.querystring("userID") %>"><i class="fa fa-angle-double-left fa-lg"></i></a></li>
<li class="page-item"><a class="page-link text-ltpurple" href="?action=<&amp;intpage=<%=intpage%>&amp;userID=<%= request.querystring("userID") %>"><i class="fa fa-angle-left fa-lg"></i></a></li>
<% end if %>
<li class="page-item"><span class="page-link text-white" style="background-color:#696986"><%=Intpage %></span></li>
<% if Intpage < intpagecount then %>
<li class="page-item"><a class="page-link text-ltpurple"  id="next-page"  href="?action=>&amp;intpage=<%=intpage%>&amp;userID=<%= request.querystring("userID") %>"><i class="fa fa-angle-right fa-lg"></i></a></li>
<li class="page-item"><a class="page-link text-ltpurple"  href="?action=>>&amp;intpage=<%=intpage%>&amp;userID=<%= request.querystring("userID") %>"><i class="fa fa-angle-double-right fa-lg mr-2"></i><%= intpagecount %></a></li>
<% end if 
 end if 
   end if ' if intpagecount <> 1
%>
</ul>
</nav>



<% 
if not rsGetProducts.eof then

rsGetProducts.AbsolutePage = intPage '======== PAGING
For intRecord = 1 To rsGetProducts.PageSize 

    ' --- pull details
    set objCmd = Server.CreateObject("ADODB.command")
    objCmd.ActiveConnection = DataConn
    objCmd.CommandText = "SELECT ID_Description, BinNumber_Detail, location, ProductDetailID, ProductDetails.ProductID as ProductID, qty, qty_counted_discontinued, item_pulled, Gauge, Length, ProductDetail1  FROM ProductDetails INNER JOIN TBL_GaugeOrder ON COALESCE (ProductDetails.Gauge, '') = COALESCE (TBL_GaugeOrder.GaugeShow, '') INNER JOIN TBL_Barcodes_SortOrder ON ProductDetails.DetailCode = TBL_Barcodes_SortOrder.ID_Number  INNER JOIN jewelry ON ProductDetails.ProductID = jewelry.ProductID WHERE ProductDetails.ProductID = ? ORDER BY ProductDetails.active DESC, ID_BarcodeOrder ASC, location ASC, item_order ASC, GaugeOrder ASC, price ASC"
    objCmd.Parameters.Append(objCmd.CreateParameter("ID",3,1,15,rsGetProducts.Fields.Item("ProductID").Value ))
    set rsGetDetails = objCmd.Execute()
%>
<button class="btn btn-primary mb-2" id="btn-finalize" id="btn-finalize" name="button" data-productid="<%= rsGetProducts.Fields.Item("ProductID").Value %>">FINALIZE</button>
<div id="message"></div>
<table class="table table-striped table-hover table-sm small">
        <thead class="thead-dark">
                <tr>
                  <th colspan="3">
                        <a href="product-edit.asp?ProductID=<%= rsGetProducts.Fields.Item("ProductID").Value %>" target="_blank"><img src="https://bafthumbs-400.bodyartforms.com/<%= rsGetProducts.Fields.Item("picture").Value %>" style="width:50px;height:auto"></a>  
                    <%= rsGetProducts.Fields.Item("title").Value %>
                    <% if ISNULL(rsGetProducts.Fields.Item("who_pulled").Value) then %>
                    <% else %>
                    <span class="badge badge-info ml-2" style="font-size:1.25em">Pulled by <%= rsGetProducts.Fields.Item("who_pulled").Value %> on <%= rsGetProducts.Fields.Item("date_pulled").Value %></span>
                    <% end if %>
                </th>
                </tr>
                <tr>
                        <th width="20%">
                          Current website stock
                      </th>
                      <th width="30%">
                            Packer counted stock
                        </th>
                        <th width="50%">
                                Item
                            </th>
                      </tr>
        </thead>
<tbody>
<% while not rsGetDetails.eof 
%>
<tr>
    <td class="ajax-update">
        <input type="text" class="form-control form-control-sm  w-25" name="qty_<%= rsGetDetails.Fields.Item("ProductDetailID").Value %>" data-column="qty"  data-id="<%= rsGetDetails.Fields.Item("ProductDetailID").Value %>" data-friendly="Product qty" value=" <%= rsGetDetails.Fields.Item("qty").Value %>">
</td>
<td>
        <%= rsGetDetails.Fields.Item("qty_counted_discontinued").Value %>

</td>
<td>
    <% If (rsGetDetails.Fields.Item("Gauge").Value) <> "" Then %>
    <%= Server.HtmlEncode(rsGetDetails.Fields.Item("Gauge").Value)%>
<% end if %>
&nbsp;&nbsp;
<% If (rsGetDetails.Fields.Item("Length").Value) <> "" Then %>
    <%= Server.HtmlEncode(rsGetDetails.Fields.Item("Length").Value)%>
<% end if %>
&nbsp;&nbsp;
<% if rsGetDetails.fields.item("ProductDetail1").value <> "" then%>
    <%= Server.HTMLEncode(rsGetDetails.Fields.Item("ProductDetail1").Value)%>
<% end if %>

</td>
</tr>
<% rsGetDetails.movenext()
    wend 
    %>
</tbody>
</table>
<%         rsGetProducts.MoveNext()
If rsGetProducts.EOF Then Exit For  ' ====== PAGING
Next ' ====== PAGING
    %>
    <nav>
        <ul class="pagination" style="justify-content: center">
    <%	if Intpage <= intpagecount then
        if intpagecount <> 1 then 
         if Intpage <> 1 then %>
    <li class="page-item"><a class="page-link text-ltpurple"  href="?action=<<&amp;intpage=<%=intpage%>&amp;userID=<%= request.querystring("userID") %>"><i class="fa fa-angle-double-left fa-lg"></i></a></li>
    <li class="page-item"><a class="page-link text-ltpurple"  href="?action=<&amp;intpage=<%=intpage%>&amp;userID=<%= request.querystring("userID") %>"><i class="fa fa-angle-left fa-lg"></i></a></li>
    <% end if %>
    <li class="page-item"><span class="page-link text-white" style="background-color:#696986"><%=Intpage %></span></li>
    <% if Intpage < intpagecount then %>
    <li class="page-item"><a class="page-link text-ltpurple"  href="?action=>&amp;intpage=<%=intpage%>&amp;userID=<%= request.querystring("userID") %>"><i class="fa fa-angle-right fa-lg"></i></a></li>
    <li class="page-item"><a class="page-link text-ltpurple"  href="?action=>>&amp;intpage=<%=intpage%>&amp;userID=<%= request.querystring("userID") %>"><i class="fa fa-angle-double-right fa-lg mr-2"></i><%= intpagecount %></a></li>
    <% end if 
     end if 
       end if ' if intpagecount <> 1
    %>
    </ul>
    </nav>
    <% end if 'not rsGetProducts.eof %>
    




<% else ' unathorized access error %>
Not accessible
<% end if ' END ACCESS TO PAGE FOR ONLY USERS WHO SHOULD BE ABLE TO SEE IT %>

</div>
<script type="text/javascript" src="../js/jquery-3.3.1.min.js"></script>
<script type="text/javascript" src="../js/bootstrap-v4.min.js"></script>
<script type="text/javascript" src="scripts/generic_auto_update_fields.js"></script>
<script type="text/javascript">
    var auto_url = "inventory/ajax-review-pulled-discontinued-items.asp"
    auto_update(); // run function to update fields when tabbing out of them
    
    // Finalize product and take off to be pulled status
	$("#btn-finalize").click(function () {
        productid = $(this).attr("data-productid");

        $.ajax({
        method: "post",
		url: "inventory/ajax-finalize-review-discontinued-items.asp",
		data: {productid: productid}
		})
		.done(function(msg) {
            $('#message').html('<div class="alert alert-success">SUCCESS</div>').show();
            var url = $("#next-page").attr("href")
            window.location = 'review-pulled-discontinued-items.asp' + url;
		})
		.fail(function(msg) {
			$('#message').html('<div class="alert alert-danger">Error processing</div>').show();
        });  
    });

</script>
</body>
</html>
<%
rsGetProducts.Close()
Set rsGetUser = Nothing
Set rsGetProducts = Nothing
%>
