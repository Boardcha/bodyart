<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/Connections/bodyartforms_sql_ADMIN.asp" -->
<!--#include virtual="/Connections/chilkat.asp" -->
<!--#include virtual="/admin/etsy/etsy-constants.asp" -->
<%
'==== PAGE HAS BEEN BOOSTRAPPED =======
bootstrapped = "yes"


set rest = Server.CreateObject("Chilkat_9_5_0.Rest")
set oauth1 = Server.CreateObject("Chilkat_9_5_0.OAuth1")

oauth1.ConsumerKey = etsy_consumer_key
oauth1.ConsumerSecret = etsy_consumer_secret
oauth1.Token = etsy_oauth_permanent_token
oauth1.TokenSecret = etsy_oauth_permanent_token_secret
oauth1.SignatureMethod = "HMAC-SHA1"
success = oauth1.GenNonce(16)

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

autoReconnect = 1
tls = 1
success = rest.Connect("openapi.etsy.com",443,tls,autoReconnect)
If (success = 0) Then
    Response.Write "<pre>" & Server.HTMLEncode( rest.LastErrorText) & "</pre>"
    Response.End
End If

%>
<html>
  <title>Manage Etsy Inventory</title>
    <body>
        <!--#include file="admin_header.asp"-->

        <div class="px-2">
            <h5 class="mt-3">Update Etsy Inventory</h5>
        <form class="form-inline">
            <input class="form-control mr-3" style="width:250px" name="keywords" type="text" placeholder="Keyword search">
            <button class="btn btn-primary" type="submit" formaction="etsy-manage-inventory.asp?page=1" formmethod="post">Search</button>
        </form>
<%
' Tell the REST object to use the OAuth1 object.
success = rest.SetAuthOAuth1(oauth1,1) 
success = rest.ClearAllQueryParams()  

items_per_page = 5
currentpage = request.querystring("page")

success = rest.AddQueryParam("limit",items_per_page)
success = rest.AddQueryParam("offset",5)
success = rest.AddQueryParam("page", request.querystring("page") )

if  request.form("keywords") <> "" then 
    success = rest.AddQueryParam("keywords", request.form("keywords") )
end if



jsonResponseText = rest.FullRequestNoBody("GET","/v2/shops/Bodyartforms/listings/active")
If (rest.LastMethodSuccess = 0) Then
    Response.Write "<pre>" & Server.HTMLEncode( rest.LastErrorText) & "</pre>"
    Response.End
End If

        
set jsonResponse = Server.CreateObject("Chilkat_9_5_0.JsonObject")
success = jsonResponse.Load(jsonResponseText)
jsonResponse.EmitCompact = 0



'Response.Write "<pre>" & Server.HTMLEncode( jsonResponse.Emit()) & "</pre>"
'Response.Write "<pre>" & Server.HTMLEncode( "Response status code: " & rest.ResponseStatusCode) & "</pre>"

total_records = jsonResponse.StringOf("count")
total_pages = total_records / items_per_page
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

<table class="table table-striped table-hover table-sm">

    <thead class="thead-dark sticky-top">
    <tr> 
      <th></th>
      <th>Etsy stock</th>
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
    ' Tell the REST object to use the OAuth1 object.
    success = rest.SetAuthOAuth1(oauth1,1)   
    success = rest.ClearAllQueryParams()

    jsonImageResponseText = rest.FullRequestNoBody("GET","/v2/listings/" & var_listing_id & "/images/")
    If (rest.LastMethodSuccess = 0) Then
        Response.Write "<pre>" & Server.HTMLEncode( rest.LastErrorText) & "</pre>"
        Response.End
    End If

        
    set jsonImageResponse = Server.CreateObject("Chilkat_9_5_0.JsonObject")
    success = jsonImageResponse.Load(jsonImageResponseText)
    jsonImageResponse.EmitCompact = 0

    'Response.Write "<pre>" & Server.HTMLEncode( jsonImageResponse.Emit()) & "</pre>"
    'Response.Write "<pre>" & Server.HTMLEncode( "Response status code: " & rest.ResponseStatusCode) & "</pre>"

    'response.write "<br>listing id " & var_listing_id

        j = 0
        count_j = jsonImageResponse.SizeOfArray("results")
        Do While j < count_j
        jsonImageResponse.J = j
            
            var_image = jsonImageResponse.StringOf("results[j].url_75x75")

        j = j + 1
        Loop
    
    '========= GET ETSY LISTING VARIATIONS ========================================
    ' Tell the REST object to use the OAuth1 object.
    success = rest.SetAuthOAuth1(oauth1,1)   
    success = rest.ClearAllQueryParams()

    jsonItemsResponseText = rest.FullRequestNoBody("GET","/v2/listings/" & var_listing_id & "/inventory/")
    If (rest.LastMethodSuccess = 0) Then
        Response.Write "<pre>" & Server.HTMLEncode( rest.LastErrorText) & "</pre>"
        Response.End
    End If

        
    set jsonItemResponse = Server.CreateObject("Chilkat_9_5_0.JsonObject")
    success = jsonItemResponse.Load(jsonItemsResponseText)
    jsonItemResponse.EmitCompact = 0

    'Response.Write "<pre>" & Server.HTMLEncode( jsonItemResponse.Emit()) & "</pre>"
    'Response.Write "<pre>" & Server.HTMLEncode( "Response status code: " & rest.ResponseStatusCode) & "</pre>"

    'response.write "<br>listing id " & var_listing_id

        j = 0
        count_j = jsonItemResponse.SizeOfArray("results.products")
        Do While j < count_j
          var_our_stock_warning = ""
          var_etsy_stock_warning = ""
          jsonItemResponse.J = j

            var_sku = jsonItemResponse.StringOf("results.products[j].sku")
            var_etsy_productid = jsonItemResponse.StringOf("results.products[j].product_id")
            var_etsy_offeringid = jsonItemResponse.StringOf("results.products[j].offerings[0].offering_id")
            var_qty = jsonItemResponse.IntOf("results.products[j].offerings[0].quantity")
            var_price = jsonItemResponse.StringOf("results.products[j].offerings[0].price.currency_formatted_raw")
            var_item = jsonItemResponse.StringOf("results.products[j].property_values[0].values")
            var_item = replace(var_item, "[""", "")
            var_item = replace(var_item, """]", "")
            var_item = replace(var_item, "\", "")

            '====== GET DETAILS FROM DATABASE =================
            set objCmd = Server.CreateObject("ADODB.Command")
            objCmd.ActiveConnection = DataConn
            objCmd.CommandText = "SELECT qty FROM ProductDetails WHERE ProductDetailID = ?"
            objCmd.Parameters.Append(objCmd.CreateParameter("detailid",3,1,20,var_sku))
            set rsGetItemInfo = objCmd.Execute()

            if rsGetItemInfo("qty") <= 0 then
                var_our_stock_warning = "bg-warning py-0 px-2 font-weight-bold m-0 rounded"
            end if
            if var_qty <= 0 then
                var_etsy_stock_warning = "bg-warning font-weight-bold"
            end if
%>

              <tr class="<% if var_qty < 5 then %>table-warning<% end if %>">
                  <td>
                    <a href="SearchDetailID.asp?DetailID=<%= var_sku %>" target="_blank"><img src="<%= var_image %>"></a>


                  </td>
                  <td class="form-inline">
                    UPDATE DOESN'T WORK YET <input class="form-control form-control-sm update-qty <%= var_etsy_stock_warning %>" style="width:75px" type="text" value="<%= var_qty %>" data-sku="<%= var_sku %>" data-listingid="<%= var_listing_id %>" data-productid="<%= var_etsy_productid %>" data-offeringid="<%= var_etsy_offeringid %>"><span id="spinner_<%= var_sku %>" style="display:none"><i class="fa fa-spinner fa-spin ml-3"></i></span>
                    <%= var_price %>
                  </td>
                  <td>
                    <span class="<%= var_our_stock_warning %>"><%= rsGetItemInfo.Fields.Item("qty").Value %></span>
                  </td>
                  <td>
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

	$(".update-qty").change(function() {
    	var listingid = $(this).attr("data-listingid");
        var productid = $(this).attr("data-productid");
        var offeringid = $(this).attr("data-offeringid");
        var sku = $(this).attr("data-sku");
        var qty = $(this).val();

        $('#spinner_' + sku).show();

		$.ajax({
		method: "post",
		url: "etsy/etsy-update-stock.asp",
        data: {listingid: listingid, productid: productid, offeringid: offeringid, sku: sku, qty: qty}
		})
		.done(function(msg) {
            $("#etsy-status").html('<div class="alert alert-success">Etsy orders have been imported</div>');
            $('#spinner_' + sku).hide();
		})
		.fail(function(msg) {
            $("#etsy-status").html('<div class="alert alert-danger">Etsy failed</div>');
            $('#spinner_' + sku).hide();
		})
	});
	
</script>
</html>
