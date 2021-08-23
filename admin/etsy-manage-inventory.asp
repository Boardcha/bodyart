<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/admin/etsy-v3/etsy-refresh-token.asp" -->
<%

'==== PAGE HAS BEEN BOOSTRAPPED =======
bootstrapped = "yes"

set rest = Server.CreateObject("Chilkat_9_5_0.Rest")

autoReconnect = 1
tls = 1
success = rest.Connect("openapi.etsy.com",443,tls,autoReconnect)
If (success = 0) Then
    Response.Write "<pre>" & Server.HTMLEncode( rest.LastErrorText) & "</pre>"
    Response.End
End If

set sbAuthHeaderVal = Server.CreateObject("Chilkat_9_5_0.StringBuilder")
success = sbAuthHeaderVal.Append("Bearer ")
success = sbAuthHeaderVal.Append(etsy_access_token)
rest.Authorization = sbAuthHeaderVal.GetAsString()
%>
<html>
<title>Manage Etsy Inventory</title>
    <body>
        <!--#include file="admin_header.asp"-->
		<link href="/CSS/fortawesome/css/external-min.css?v=031920" rel="stylesheet" type="text/css" />
        <div class="px-2">
            <h5 class="mt-3 mb-3">Manage Etsy Inventory</h5>
        <form class="form-inline">
            <input class="form-control mr-3" style="width:250px" name="keywords" type="text" placeholder="Keyword search" value="<%= request.form("keywords")%>">
            <button class="btn btn-primary" type="submit" formaction="etsy-manage-inventory.asp?page=1" formmethod="post">Search</button>
        </form>
		<%If request.form("keywords")<>"" Then %>
		<div>
			<a class="filter-delete text-danger d-lg-block d-inline-block mr-3 mr-lg-0" href="/admin/etsy-manage-inventory.asp?page=1" data-filter="keywords">
			<i class="fa fa-times"></i>
			<%= request.form("keywords")%>
			</a>
		</div>	
		<%End If%>
<%

success = rest.ClearAllQueryParams()  

items_per_page = 5
currentpage = request.querystring("page")
If IsNumeric(currentpage) Then
	If currentpage > 0 Then offset = (currentpage - 1) * 5 Else offset = 0
Else 
	offset = 0
End If
success = rest.AddQueryParam("client_id", etsy_consumer_key)
success = rest.AddQueryParam("limit", items_per_page)
success = rest.AddQueryParam("offset", offset)

if  request.form("keywords") <> "" then 
    success = rest.AddQueryParam("keywords", request.form("keywords") )
end if


jsonResponseText = rest.FullRequestNoBody("GET","/v3/application/shops/" & etsy_baf_shop_id & "/listings/active")
If (rest.LastMethodSuccess = 0) Then
    Response.Write "<pre>" & Server.HTMLEncode( rest.LastErrorText) & "</pre>"
    Response.End
End If

        
set jsonResponse = Server.CreateObject("Chilkat_9_5_0.JsonObject")
success = jsonResponse.Load(jsonResponseText)
jsonResponse.EmitCompact = 0

'Response.Write "<pre>" & Server.HTMLEncode( jsonResponse.Emit()) & "</pre>"
'Response.Write "<pre>" & Server.HTMLEncode( "Response status code: " & rest.ResponseStatusCode) & "</pre>"
'Response.End


total_records = CLng(jsonResponse.StringOf("count"))
total_pages = Int(total_records / items_per_page)
'Response.Write "total_records:" & total_records & "<br>"
'Response.Write "total_pages:" & total_pages & "<br>"
'Response.Write "offset:" & offset & "<br>"
%>

<nav aria-label="Page navigation example">
    <ul class="pagination justify-content-center">
        <li class="page-item">
            <a class="page-link" href="etsy-manage-inventory.asp?page=1">First</a>
          </li>
          <% if cint(currentpage) > 2 then %>
      <li class="page-item"><a class="page-link" href="etsy-manage-inventory.asp?page=<%= currentpage -2 %>"><%= currentpage - 2 %></a></li>
      <li class="page-item"><a class="page-link" href="etsy-manage-inventory.asp?page=<%= currentpage -1 %>"><%= currentpage - 1 %></a></li>
      <% end if %>
      <li class="page-item active"><a class="page-link" href="etsy-manage-inventory.asp?page=<%= currentpage %>"><%= currentpage %></a></li>

      <% if cint(currentpage) < cint(total_pages) then %>
      <li class="page-item"><a class="page-link" href="etsy-manage-inventory.asp?page=<%= currentpage + 1 %>"><%= currentpage + 1 %></a></li>
      <li class="page-item"><a class="page-link" href="etsy-manage-inventory.asp?page=<%= currentpage + 2 %>"><%= currentpage + 2 %></a></li>
      <% end if %>
      <li class="page-item">
        <a class="page-link" href="etsy-manage-inventory.asp?page=<%= total_pages %>">Last | <%= total_pages %></a>
      </li>
    </ul>
</nav>

  
<table class="table table-striped table-sm">

    <thead class="thead-dark sticky-top">
    <tr> 
      <th></th>
      <th>Etsy stock</th>
	  <th>Etsy price</th>
      <th>Our current stock</th>
      <th>Etsy description</th>
    </tr>
  </thead>

<%
i = 0
count_i = jsonResponse.SizeOfArray("results")
Do While i < count_i
	jsonResponse.I = i

	var_listing_id = jsonResponse.StringOf("results[i].listing_id")
	var_title = jsonResponse.StringOf("results[i].title")


	var_title = CleanUp(var_title)
	var_title = replace(var_title, "g ", "")
	var_title = replace(var_title, "mm", "")
	var_title = replace(var_title, "quot", "")

    '=========== GET ETSY LISTING IMAGE =================================
 
    success = rest.ClearAllQueryParams()
	success = rest.AddQueryParam("client_id", etsy_consumer_key)

    jsonImageResponseText = rest.FullRequestNoBody("GET","/v3/application/shops/" & etsy_baf_shop_id & "/listings/" & var_listing_id & "/images")
    If (rest.LastMethodSuccess = 0) Then
        Response.Write "<pre>" & Server.HTMLEncode( rest.LastErrorText) & "</pre>"
        Response.End
    End If

        
    set jsonImageResponse = Server.CreateObject("Chilkat_9_5_0.JsonObject")
    success = jsonImageResponse.Load(jsonImageResponseText)
    jsonImageResponse.EmitCompact = 0

    'Response.Write "<pre>" & Server.HTMLEncode( jsonImageResponse.Emit()) & "</pre>"
    'Response.Write "<pre>" & Server.HTMLEncode( "Response status code: " & rest.ResponseStatusCode) & "</pre>"
	'Response.End
	
    'response.write "<br>listing id " & var_listing_id

	j = 0
	count_j = jsonImageResponse.SizeOfArray("results")
	Do While j < count_j
		jsonImageResponse.J = j
		var_image = jsonImageResponse.StringOf("results[j].url_75x75")
		j = j + 1
	Loop
    
    '========= GET ETSY LISTING VARIATIONS ========================================
    success = rest.ClearAllQueryParams()
	success = rest.AddQueryParam("client_id", etsy_consumer_key)

    jsonItemsResponseText = rest.FullRequestNoBody("GET","/v3/application/listings/" & var_listing_id & "/inventory")
    If (rest.LastMethodSuccess = 0) Then
        Response.Write "<pre>" & Server.HTMLEncode( rest.LastErrorText) & "</pre>"
        Response.End
    End If

        
    set jsonItemResponse = Server.CreateObject("Chilkat_9_5_0.JsonObject")
    success = jsonItemResponse.Load(jsonItemsResponseText)
    jsonItemResponse.EmitCompact = 0

    'Response.Write "<pre>" & Server.HTMLEncode( jsonItemResponse.Emit()) & "</pre>"
    'Response.Write "<pre>" & Server.HTMLEncode( "Response status code: " & rest.ResponseStatusCode) & "</pre>"
	'Response.End
    'response.write "<br>listing id " & var_listing_id

        j = 0
        count_j = jsonItemResponse.SizeOfArray("products")
		
        Do While j < count_j
            jsonItemResponse.J = j
			
            var_sku = jsonItemResponse.StringOf("products[j].sku")
            var_etsy_productid = jsonItemResponse.StringOf("products[j].product_id")
            var_etsy_offeringid = jsonItemResponse.StringOf("products[j].offerings[0].offering_id")
            var_qty = jsonItemResponse.IntOf("products[j].offerings[0].quantity")
			var_etsy_is_enabled = jsonItemResponse.IntOf("products[j].offerings[0].is_enabled")
            var_price = jsonItemResponse.StringOf("products[j].offerings[0].price.amount")
			if jsonItemResponse.IntOf("products[j].offerings[0].price.divisor") > 0 Then var_price = jsonItemResponse.StringOf("products[j].offerings[0].price.amount") / jsonItemResponse.IntOf("products[j].offerings[0].price.divisor")
			var_etsy_propertyid = jsonItemResponse.StringOf("products[j].property_values[0].property_id")
			var_etsy_scaleid = jsonItemResponse.StringOf("products[j].property_values[0].scale_id")
			var_etsy_property_name = jsonItemResponse.StringOf("products[j].property_values[0].property_name")
			var_etsy_property_valueId = jsonItemResponse.StringOf("products[j].property_values[0].value_ids[0]")
			var_etsy_property_values = jsonItemResponse.StringOf("products[j].property_values[0].values[0]")
            var_item = jsonItemResponse.StringOf("products[j].property_values[0].values")
            var_item = replace(var_item, "[""", "")
            var_item = replace(var_item, """]", "")
            var_item = replace(var_item, "\", "")

            '====== GET DETAILS FROM DATABASE =================
            set objCmd = Server.CreateObject("ADODB.Command")
            objCmd.ActiveConnection = DataConn
            objCmd.CommandText = "SELECT qty FROM ProductDetails WHERE ProductDetailID = ?"
            objCmd.Parameters.Append(objCmd.CreateParameter("detailid",3,1,20,var_sku))
            set rsGetItemInfo = objCmd.Execute()
%>
              <tr class="<% if var_qty < 5 then %>table-warning<% end if %>">
                  <td class="align-middle">
                    <a href="SearchDetailID.asp?DetailID=<%= var_sku %>" target="_blank"><img src="<%= var_image %>"></a>
                  </td>
                  <td class="align-middle">
                    <input class="form-control form-control-sm update-qty d-inline-block" style="width:75px" type="text" value="<%= var_qty %>" data-price="<%= Replace(var_price, ",",".") %>" data-listingid="<%= var_listing_id %>" data-sku="<%= var_sku %>" data-productid="<%= var_etsy_productid %>"><div class="spinner" style="display:none"><i class="fa fa-spinner fa-spin ml-3"></i></div>
                  </td>
                  <td class="align-middle">
                    <input class="form-control form-control-sm update-price d-inline-block" style="width:75px" type="text" value="<%= Replace(var_price, ",",".") %>" data-price="<%= Replace(var_price, ",",".") %>" data-listingid="<%= var_listing_id %>" data-sku="<%= var_sku %>" data-productid="<%= var_etsy_productid %>"><div class="spinner" style="display:none"><i class="fa fa-spinner fa-spin ml-3"></i></div>
                  </td>				  
                  <td class="align-middle" style="padding-left:60px">
                    <%= rsGetItemInfo.Fields.Item("qty").Value %>
                  </td>
                  <td class="align-middle">
                    <span class="mr-3"><%= var_title %></span>
                     <%= var_item %>
                  </td>
              </tr>
<%
            Set rsGetItemInfo = Nothing
            j = j + 1
        Loop


    i = i + 1
Loop

%>   
</table> 

<nav class="mt-4" aria-label="Page navigation example">
    <ul class="pagination justify-content-center">
        <li class="page-item">
            <a class="page-link" href="etsy-manage-inventory.asp?page=1">First</a>
          </li>
          <% if cint(currentpage) > 2 then %>
      <li class="page-item"><a class="page-link" href="etsy-manage-inventory.asp?page=<%= currentpage -2 %>"><%= currentpage - 2 %></a></li>
      <li class="page-item"><a class="page-link" href="etsy-manage-inventory.asp?page=<%= currentpage -1 %>"><%= currentpage - 1 %></a></li>
      <% end if %>
      <li class="page-item active"><a class="page-link" href="etsy-manage-inventory.asp?page=<%= currentpage %>"><%= currentpage %></a></li>

      <% if cint(currentpage) < cint(total_pages) then %>
      <li class="page-item"><a class="page-link" href="etsy-manage-inventory.asp?page=<%= currentpage + 1 %>"><%= currentpage + 1 %></a></li>
      <li class="page-item"><a class="page-link" href="etsy-manage-inventory.asp?page=<%= currentpage + 2 %>"><%= currentpage + 2 %></a></li>
      <% end if %>
      <li class="page-item">
        <a class="page-link" href="etsy-manage-inventory.asp?page=<%= total_pages %>">Last | <%= total_pages %></a>
      </li>
    </ul>
  </nav>
</div>
</body>
<script src="/js/jquery-3.3.1.min.js"></script>
<script type="text/javascript">

	$(".update-qty, .update-price").change(function() {
    	var listingid = $(this).attr("data-listingid");
        var productid = $(this).attr("data-productid");
		var sku = $(this).attr("data-sku");
        var qty = $(this).parent().parent().find(".update-qty").val();
		var price = $(this).parent().parent().find(".update-price").val();
        
		var element = $(this);
		element.siblings(".spinner").css("display", "inline-block");
		
		$.ajax({
		method: "post",
		url: "etsy-v3/etsy-update-stock.asp",
        data: {listingid: listingid, productid: productid, qty: qty, price: price, sku: sku}
		})
		.done(function(msg) {
            $("#etsy-status").html('<div class="alert alert-success">Etsy orders have been imported</div>');
            element.siblings(".spinner").hide();
			element.closest("tr").removeClass("table-warning");
			element.closest("tr").removeClass("table-danger");	
			element.closest("tr").addClass("table-info");	
		})
		.fail(function(msg) {
            $("#etsy-status").html('<div class="alert alert-danger">Etsy failed</div>');
            element.siblings(".spinner").hide();
			element.closest("tr").addClass("table-danger");	
		})
	});
	
</script>
</html>
<%
Set rest = Nothing
Set sbAuthHeaderVal = Nothing
Set jsonResponse = Nothing
Set jsonImageResponse = Nothing
Set objCmd = Nothing
Set rsGetItemInfo = Nothing
DataConn.Close

Function CleanUp (input)
    Dim objRegExp, outputStr
    Set objRegExp = New Regexp

    objRegExp.IgnoreCase = True
    objRegExp.Global = True
    objRegExp.Pattern = "((?![a-zA-Z]).)+"
    outputStr = objRegExp.Replace(input, "-")

    objRegExp.Pattern = "\-+"
    outputStr = objRegExp.Replace(outputStr, " ")

    CleanUp = outputStr
End Function
%>