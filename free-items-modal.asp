<!--#include virtual="/Connections/sql_connection.asp" -->
<%
	
page_title = "Free Items"	
title_onpage = "Free Items"
page_description = "Free Items"
var_photo_url = "https://bafthumbs-400.bodyartforms.com"

SQL = _
"SELECT * FROM(" & _
	"SELECT jewelry.ProductID, jewelry.title, ProductDetails.ProductDetailID, ProductDetails.price, ProductDetails.qty, ProductDetails.free_item_start_date, ProductDetails.free_item_expiration_date, ProductDetails.ProductDetail1, ProductDetails.Gauge, jewelry.picture, ProductDetails.detail_code, ProductDetails.Free_QTY, ProductDetails.free, " & _
	"ROW_NUMBER() OVER (PARTITION BY jewelry.ProductID, ProductDetails.free ORDER BY jewelry.ProductID DESC) AS [ROWNUMBER], " & _
	"Count(jewelry.ProductID) OVER (PARTITION BY jewelry.ProductID) AS [ROWCOUNT] " & _
	"FROM ProductDetails INNER JOIN jewelry ON ProductDetails.ProductID = jewelry.ProductID " & _
	"WHERE (jewelry.ProductID <> 3704) AND (ProductDetails.qty > 0) AND (ProductDetails.free <> 0) AND (free_item_start_date <= GETDATE() OR free_item_expiration_date is Null) AND (free_item_expiration_date > GETDATE() OR free_item_expiration_date is Null) AND (ProductDetails.free IS NOT NULL) AND (ProductDetails.active = 1)) GROUPS " & _
"WHERE GROUPS.[ROWNUMBER] = 1 ORDER BY GROUPS.free, GROUPS.ProductDetailID"
	  
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = SQL
Set rsGetRecords = objCmd.Execute()
%>


<div class="products">

	Use this free items guide to help you plan your shopping! When you meet certain cart thresholds you will see eligible free items to select from on the shopping cart page right before you checkout.
	<% if NOT rsGetRecords.EOF then	%>
		<div class="d-flex flex-row flex-wrap">
			<% While NOT rsGetRecords.EOF %>
				<%If treshold <> rsGetRecords("free") Then%>
				<div class="alert alert-secondary w-100 mb-0 mt-3">
					<div class="h4">
						<i class="fa fa-lg fa-chevron-down mr-3"></i><%= "$" & FormatNumber(rsGetRecords("free"), 2) & " CART THRESHOLD"%>
					</div>
					<% if rsGetRecords("free") > 30 then %>
						Select one item from the previous threshold selections AND get one item from this new threshold.
						<% else %>
						Select one item from this first threshold
					<% end if %>
				</div>
				<%End If%>
				
				<% treshold = rsGetRecords("free") %>
				<div class="col-12 col-xs-4 col-md-3 col-xl-3 col-break1600-2 my-3 px-1 px-md-2 text-center">	
					<div class="container-fluid">
						<div class="row">	
								<div class="position-relative">
									
									<img class="img-fluid w-100 <% if lazy_count > 20 then %> lazyload <% end if %>" <% if lazy_count > 20 then %> src="/images/image-placeholder.png" data-src="<%= var_photo_url %>/<%=(rsGetRecords("picture"))%>" <% else %> src="<%= var_photo_url %>/<%=(rsGetRecords("picture"))%>" <% end if %> title="<%=(rsGetRecords("title"))%>" alt="<%=(rsGetRecords("title"))%>" />

								</div>
						</div>					
							<div class="row text-center">
								<div class="w-100 h6">
									<%If rsGetRecords("title") = "Order Credit" Then %>
										<%= "$" & FormatNumber(rsGetRecords("price")) & " " & rsGetRecords("title")%>
									<%Else%>
										<%=rsGetRecords("title")%>
									<%End If%>	
								</div> 
								<%If rsGetRecords("rowcount") > 1 Then %>
									<div class="w-100 px-1">
										<%If rsGetRecords("title") <> "Order Credit" Then %>
											<button class="btn btn-sm btn-outline-secondary view-variations" type="button" data-id="<%= rsGetRecords("ProductID") %>" data-toggle="modal" data-target="#free-items-page-variations-modal">View <%=rsGetRecords("rowcount")%> available variations</button>
										<% end if %>
									</div>
								<%End If%>	
							</div>
					</div>
				</div>
				<%
				lazy_count = lazy_count + 1
				rsGetRecords.MoveNext()
			Wend 
			%>
		</div>
	<%Else%>
		<h5 class="alert alert-danger mt-3">No results found</h5>
	<% End If%>
</div>

