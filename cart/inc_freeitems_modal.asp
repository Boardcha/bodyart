<!--#include virtual="/template/inc_includes_ajax.asp" -->
<%
free_items_count = Request("free_items_count")
var_viewcart_showgifts = "yes"

if request.cookies("gaugecard") <> "no" then
	'Make sure it is added by default if there is no cookie
	Response.Cookies("gaugecard")= "yes"
end if
%>

<div class="custom-control custom-checkbox form-group form-check form-check-inline pl-0">
	<div class="form-check-label-gauge d-inline-block"></div>
	<input type="checkbox" class="custom-control-input form-check-input" id="gaugeCardCheck" <%if request.cookies("gaugecard") <> "no" then %> checked<%end if%>>
	<label class="custom-control-label mt-3" for="gaugeCardCheck">Gauge card</label>  
</div>				

<!--Accordion wrapper-->
<div class="accordion md-accordion" id="accordionEx" role="tablist" aria-multiselectable="true">
<%
' ------- Get FREE items for TIER 1 (ORINGS)
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "SELECT DISTINCT FlatProducts.min_gauge, FlatProducts.max_gauge, jewelry.title, jewelry.picture,jewelry.ProductID, jewelry.picture, CASE WHEN jewelry.ProductID = 2890 THEN '1' ELSE jewelry.ProductID END " & _
							"FROM ProductDetails INNER JOIN jewelry ON ProductDetails.ProductID = jewelry.ProductID " & _ 
							"INNER JOIN FlatProducts ON FlatProducts.ProductID = jewelry.ProductID " & _ 
							"WHERE ProductDetails.free <= 30 AND (free_item_expiration_date > GETDATE() OR free_item_expiration_date is Null) AND (jewelry.ProductID = 530 OR (jewelry.ProductID = 1649 AND ProductDetails.Gauge <> '1-1/8" & """" & "') OR jewelry.ProductID = 15385) " & _
							"ORDER BY CASE WHEN jewelry.ProductID = 2890 THEN '1' ELSE jewelry.ProductID END, jewelry.title"
		Set rsGetFreeTier1 = objCmd.Execute()
' ------- End getting free items for TIER 1
%>		
		  <!-- Accordion card -->
		  <div class="card" style="overflow: visible;">

			<!-- Card header -->
			<div class="card-header" role="tab" id="headingOne1">
			  <a data-toggle="collapse" data-parent="#accordionEx" href="#collapseOne1" aria-expanded="true"
				aria-controls="collapseOne1" class="accordion-button">
				<div class="d-inline-block">
				<h5 class="mb-0">
				  FREE O-RING <i class="fas fa-angle-down rotate-icon"></i>
				</h5>
				</div>
				<div id="header-card-selected-item-1"></div>				
			  </a>
			</div>

  
			<!-- Card body -->
			<div id="collapseOne1" class="collapse show" data-tier="1" role="tabpanel" aria-labelledby="headingOne1" data-parent="#accordionEx">
			  <div class="card-body">
				<div class="mt-2 full-width-block">

					<div class="select-product-variation" id="select-product-variation-1" style="z-index: 2000; position: absolute; width: 80%; top: 117px; left: 0; right: 0; margin: auto; padding:0 !important" class="modal-body"></div>

					<div class="baf-carousel slick-free-items" id="slick-free-items-1"  style="">
						<% rsGetFreeTier1.MoveFirst %>	
						<% index = 0 %>
						<% Do While NOT rsGetFreeTier1.EOF %>	
							<% index = index + 1 %>
							<a class="slide text-dark homepage-graphic" data-slide-index="<%=index%>" data-tier="1" data-productid="<%= rsGetFreeTier1("ProductID") %>" id="tier-1-<%=index%>">
								<img class="img-fluid" src="https://bafthumbs-400.bodyartforms.com/<%= rsGetFreeTier1("picture") %>" alt="<%=(rsGetFreeTier1("title"))%>" title="<%=(rsGetFreeTier1("title"))%>">
								<div class="w-100 text-center px-1 variation-text">
									<div class="small">
											<%= rsGetFreeTier1("title") %>
									</div>
									<div class="small font-weight-bold  d-block">
											<%= rsGetFreeTier1("min_gauge") %>
											<% if rsGetFreeTier1("min_gauge") <> rsGetFreeTier1("max_gauge") then %> 
											- <%= rsGetFreeTier1("max_gauge") %>
											<% end if %>
									</div>
								</div> 
								<div class="w-100 text-center px-1 selected-variation-text" style="display:none">
								</div>
							</a>
						<% rsGetFreeTier1.MoveNext()
						Loop
						%>
					</div>		
				</div>			
			  </div>
			</div>

		  </div>
		  <!-- Accordion card -->

<%
' ------- Get FREE items for TIER 2 (STICKERS)
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "SELECT ProductDetails.detail_code, ProductDetails.free, 1 As Free_QTY, jewelry.picture, ProductDetails.ProductDetailID, ProductDetails.ProductDetail1, FlatProducts.min_gauge, FlatProducts.max_gauge, jewelry.title, jewelry.picture,jewelry.ProductID, jewelry.picture, CASE WHEN jewelry.ProductID = 2890 THEN '1' ELSE jewelry.ProductID END, ISNULL(ProductDetails.gauge,'') + ' ' + ISNULL(ProductDetails.Length,'') + ' ' + ISNULL(ProductDetails.ProductDetail1,'') + ' ' + ISNULL(jewelry.title,'') AS 'free_title' " & _
						"FROM ProductDetails INNER JOIN jewelry ON ProductDetails.ProductID = jewelry.ProductID " & _ 
						"INNER JOIN FlatProducts ON FlatProducts.ProductID = jewelry.ProductID " & _ 
						"INNER JOIN TBL_GaugeOrder Gauge ON ISNULL(ProductDetails.Gauge,'') = ISNULL(Gauge.GaugeShow,'') " & _ 
						"WHERE (jewelry.ProductID = 3928) AND (ProductDetails.qty > 0) AND (ProductDetails.free <> 0) AND (ProductDetails.free IS NOT NULL) AND (ProductDetails.free_item_expiration_date > GETDATE() OR ProductDetails.free_item_expiration_date is Null) AND (ProductDetails.active = 1) " & _
						"ORDER BY GaugeOrder ASC, item_order ASC, Price ASC"
		Set rsGetFreeTier2 = objCmd.Execute()
' ------- End getting free items for TIER 2
%>			  
		  <!-- Accordion card -->
		  <div class="card" style="overflow: visible;">

			<!-- Card header -->
			<div class="card-header" role="tab" id="headingOne2">
			  <a data-toggle="collapse" data-parent="#accordionEx" href="#collapseOne2" aria-expanded="false"
				aria-controls="collapseOne2" class="accordion-button collapsed">
				<div class="d-inline-block">
				<h5 class="mb-0">
				  FREE STICKER <i class="fas fa-angle-down rotate-icon"></i>
				</h5>
				</div>
				<div id="header-card-selected-item-2"></div>				
			  </a>
			</div>

			<!-- Card body -->
			<div id="collapseOne2" class="collapse" data-tier="2" role="tabpanel" aria-labelledby="headingOne2" data-parent="#accordionEx">
			  <div class="card-body">
				<div class="mt-2 full-width-block">

					<div class="select-product-variation" id="select-product-variation-2" style="z-index: 2000; position: absolute; width: 80%; top: 117px; left: 0; right: 0; margin: auto; padding:0 !important" class="modal-body"></div>

					<div class="baf-carousel slick-free-items free-stickers" id="slick-free-items-2"  style="">
						<% rsGetFreeTier2.MoveFirst %>	
						<% index = 0 %>
						<% Do While NOT rsGetFreeTier2.EOF %>	
							<% index = index + 1 %>
							<a class="slide text-dark homepage-graphic" data-slide-index="<%=index%>" data-tier="2" data-product-detail-id="<%= rsGetFreeTier2("ProductDetailID") %>" data-friendly="<%= rsGetFreeTier2("ProductDetail1") %>" id="tier-2-<%=index%>">
								<img width="138" class="img-fluid" src="https://bafthumbs-400.bodyartforms.com/<%= rsGetFreeTier2("detail_code") %>" alt="<%=(rsGetFreeTier2("title"))%>" title="<%=(rsGetFreeTier2("title"))%>">
								<div class="w-100 text-center px-1 variation-text">
									<div class="small">
											<%= rsGetFreeTier2("ProductDetail1") %>
									</div>
									<div class="small font-weight-bold  d-block">
											<%= rsGetFreeTier2("min_gauge") %>
											<% if rsGetFreeTier2("min_gauge") <> rsGetFreeTier2("max_gauge") then %> 
											- <%= rsGetFreeTier2("max_gauge") %>
											<% end if %>
									</div>
								</div> 
								<div class="w-100 text-center px-1 selected-variation-text" style="display:none">
								</div>
							</a>
						<% rsGetFreeTier2.MoveNext()
						Loop
						%>
					</div>		
				</div>			
			  </div>
			</div>
		  </div>
		  <!-- Accordion card -->

<%
' ------- Get FREE items for TIER 3
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "SELECT DISTINCT FlatProducts.min_gauge, FlatProducts.max_gauge, jewelry.title, jewelry.picture,jewelry.ProductID, jewelry.picture, CASE WHEN jewelry.ProductID = 2890 THEN '1' ELSE jewelry.ProductID END " & _
							"FROM ProductDetails INNER JOIN jewelry ON ProductDetails.ProductID = jewelry.ProductID " & _ 
							"INNER JOIN FlatProducts ON FlatProducts.ProductID = jewelry.ProductID " & _ 
							"WHERE ProductDetails.free <= 30 AND (jewelry.ProductID <> 3704) AND (ProductDetails.qty > 0) AND (ProductDetails.free <> 0) AND (ProductDetails.free IS NOT NULL) AND (ProductDetails.free_item_expiration_date > GETDATE() OR ProductDetails.free_item_expiration_date is Null) AND (ProductDetails.free IS NOT NULL) AND (ProductDetails.active = 1) " & _
							"ORDER BY CASE WHEN jewelry.ProductID = 2890 THEN '1' ELSE jewelry.ProductID END, jewelry.title"
		Set rsGetFreeTier3 = objCmd.Execute()
' ------- End getting free items for TIER 3
%>
		  <!-- Accordion card -->
		  <div class="card" style="overflow: visible;">

			<!-- Card header -->
			<div class="card-header" role="tab" id="headingThree3">
			  <a data-toggle="collapse" data-parent="#accordionEx" href="#collapseThree3"
				aria-expanded="false" aria-controls="collapseThree3" class="accordion-button collapsed">
				<div class="d-inline-block">
				<h5 class="mb-0">
				 <%If free_items_count < 3 Then %><i class="fa fa-lock"></i><%End If%> $30.00 CART VALUE <i class="fas fa-angle-down rotate-icon"></i>
				</h5>
				</div>
				<div id="header-card-selected-item-3"></div>
			  </a>
			</div>

			<!-- Card body -->
			<div id="collapseThree3" class="collapse" data-tier="3" role="tabpanel" aria-labelledby="headingThree3" data-parent="#accordionEx">
			  <div class="card-body">
				<div class="mt-2 full-width-block">
					<div class="select-product-variation" id="select-product-variation-3" style="z-index: 2000; position: absolute; width: 80%; top: 117px; left: 0; right: 0; margin: auto; padding:0 !important" class="modal-body"></div>
					<div class="baf-carousel slick-free-items <%If free_items_count < 3 Then %>notavailable<%End If%>" id="slick-free-items-3" style="">
						<% rsGetFreeTier3.MoveFirst %>	
						<% index = 0 %>
						<% Do While NOT rsGetFreeTier3.EOF %>	
							<% index = index + 1 %>
							<a class="slide text-dark homepage-graphic" data-tier="3" data-slide-index="<%=index%>" data-productid="<%= rsGetFreeTier3("ProductID") %>" id="tier-3-<%=index%>">
								<img class="img-fluid" src="https://bafthumbs-400.bodyartforms.com/<%= rsGetFreeTier3("picture") %>" alt="<%=(rsGetFreeTier3("title"))%>" title="<%=(rsGetFreeTier3("title"))%>">
								<div class="w-100 text-center px-1 variation-text">
									<div class="small">
											<%= rsGetFreeTier3("title") %>
									</div>
									<div class="small font-weight-bold  d-block">
											<%= rsGetFreeTier3("min_gauge") %>
											<% if rsGetFreeTier3("min_gauge") <> rsGetFreeTier3("max_gauge") then %> 
											- <%= rsGetFreeTier3("max_gauge") %>
											<% end if %>
									</div>
								</div> 
								<div class="w-100 text-center px-1 selected-variation-text" style="display:none">
								</div>
							</a>
						<% rsGetFreeTier3.MoveNext()
						Loop
						%>
					</div>		
				</div>
			  </div>
			</div>

		  </div>
		  <!-- Accordion card -->

<%
' ------- Get FREE items for TIER 4
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "SELECT DISTINCT FlatProducts.min_gauge, FlatProducts.max_gauge, jewelry.title, jewelry.picture,jewelry.ProductID, jewelry.picture, CASE WHEN jewelry.ProductID = 2890 THEN '1' ELSE jewelry.ProductID END " & _
							"FROM ProductDetails INNER JOIN jewelry ON ProductDetails.ProductID = jewelry.ProductID " & _ 
							"INNER JOIN FlatProducts ON FlatProducts.ProductID = jewelry.ProductID " & _ 
							"WHERE ProductDetails.free <= 50 AND (jewelry.ProductID <> 3704) AND (ProductDetails.qty > 0) AND (ProductDetails.free <> 0) AND (ProductDetails.free IS NOT NULL) AND (ProductDetails.free_item_expiration_date > GETDATE() OR ProductDetails.free_item_expiration_date is Null) AND (ProductDetails.free IS NOT NULL) AND (ProductDetails.active = 1) " & _
							"ORDER BY CASE WHEN jewelry.ProductID = 2890 THEN '1' ELSE jewelry.ProductID END, jewelry.title"
		Set rsGetFreeTier4 = objCmd.Execute()
' ------- End getting free items for TIER 4
%>
		  
		  <!-- Accordion card -->
		  <div class="card" style="overflow: visible;">

			<!-- Card header -->
			<div class="card-header" role="tab" id="headingThree4">
			  <a data-toggle="collapse" data-parent="#accordionEx" href="#collapseThree4"
				aria-expanded="false" aria-controls="collapseThree4" class="accordion-button collapsed">
				<div class="d-inline-block">
				<h5 class="mb-0">
				 <%If free_items_count  < 4 Then %><i class="fa fa-lock"></i><%End If%> $50.00 CART VALUE <i class="fas fa-angle-down rotate-icon"></i> 
				</h5>
				</div>
				<div id="header-card-selected-item-4"></div>				
			  </a>
			</div>

			<!-- Card body -->
			<div id="collapseThree4" class="collapse" data-tier="4" role="tabpanel" aria-labelledby="headingThree4" data-parent="#accordionEx">
			  <div class="card-body">
				<div class="mt-2 full-width-block">
					<div class="select-product-variation" id="select-product-variation-4" style="z-index: 2000; position: absolute; width: 80%; top: 117px; left: 0; right: 0; margin: auto; padding:0 !important" class="modal-body"></div>
					<div class="baf-carousel slick-free-items <%If free_items_count < 4 Then %>notavailable<%End If%>" id="slick-free-items-4" style="">
						<% rsGetFreeTier4.MoveFirst %>	
						<% index = 0 %>
						<% Do While NOT rsGetFreeTier4.EOF %>	
							<% index = index + 1 %>
							<a class="slide text-dark homepage-graphic" data-tier="4" data-slide-index="<%=index%>" data-productid="<%= rsGetFreeTier4("ProductID") %>" id="tier-4-<%=index%>">
								<img class="img-fluid" src="https://bafthumbs-400.bodyartforms.com/<%= rsGetFreeTier4("picture") %>" alt="<%=(rsGetFreeTier4("title"))%>" title="<%=(rsGetFreeTier4("title"))%>">
								<div class="w-100 text-center px-1 variation-text">
									<div class="small">
											<%= rsGetFreeTier4("title") %>
									</div>
									<div class="small font-weight-bold  d-block">
											<%= rsGetFreeTier4("min_gauge") %>
											<% if rsGetFreeTier4("min_gauge") <> rsGetFreeTier4("max_gauge") then %> 
											- <%= rsGetFreeTier4("max_gauge") %>
											<% end if %>
									</div>
								</div> 
								<div class="w-100 text-center px-1 selected-variation-text" style="display:none">
								</div>
							</a>
						<% rsGetFreeTier4.MoveNext()
						Loop
						%>
					</div>		
				</div>
			  </div>
			</div>

		  </div>
		  <!-- Accordion card -->

<%
' ------- Get FREE items for TIER 5
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "SELECT DISTINCT FlatProducts.min_gauge, FlatProducts.max_gauge, jewelry.title, jewelry.picture,jewelry.ProductID, jewelry.picture, CASE WHEN jewelry.ProductID = 2890 THEN '1' ELSE jewelry.ProductID END " & _
							"FROM ProductDetails INNER JOIN jewelry ON ProductDetails.ProductID = jewelry.ProductID " & _ 
							"INNER JOIN FlatProducts ON FlatProducts.ProductID = jewelry.ProductID " & _ 
							"WHERE ProductDetails.free <= 75 AND (jewelry.ProductID <> 3704) AND (ProductDetails.qty > 0) AND (ProductDetails.free <> 0) AND (ProductDetails.free IS NOT NULL) AND (ProductDetails.free_item_expiration_date > GETDATE() OR ProductDetails.free_item_expiration_date is Null) AND (ProductDetails.free IS NOT NULL) AND (ProductDetails.active = 1) " & _
							"ORDER BY CASE WHEN jewelry.ProductID = 2890 THEN '1' ELSE jewelry.ProductID END, jewelry.title"
		Set rsGetFreeTier5 = objCmd.Execute()
' ------- End getting free items for TIER 5
%>
		  
		  <!-- Accordion card -->
		  <div class="card" style="overflow: visible;">

			<!-- Card header -->
			<div class="card-header" role="tab" id="headingThree5">
			  <a data-toggle="collapse" data-parent="#accordionEx" href="#collapseThree5"
				aria-expanded="false" aria-controls="collapseThree5" class="accordion-button collapsed">
				<div class="d-inline-block">
				<h5 class="mb-0">
				  <%If free_items_count < 5 Then %><i class="fa fa-lock"></i><%End If%> $75.00 CART VALUE <i class="fas fa-angle-down rotate-icon"></i>
				</h5>
				</div>
				<div id="header-card-selected-item-5"></div>				
			  </a>
			</div>

			<!-- Card body -->
			<div id="collapseThree5" class="collapse" data-tier="5" role="tabpanel" aria-labelledby="headingThree5" data-parent="#accordionEx">
			  <div class="card-body">
				<div class="mt-2 full-width-block">
					<div class="select-product-variation" id="select-product-variation-5" style="z-index: 2000; position: absolute; width: 80%; top: 117px; left: 0; right: 0; margin: auto; padding:0 !important" class="modal-body"></div>
					<div class="baf-carousel slick-free-items <%If free_items_count < 5 Then %>notavailable<%End If%>" id="slick-free-items-5" style="">
						<% rsGetFreeTier5.MoveFirst %>	
						<% index = 0 %>
						<% Do While NOT rsGetFreeTier5.EOF %>	
							<% index = index + 1 %>
							<a class="slide text-dark homepage-graphic" data-tier="5" data-slide-index="<%=index%>" data-productid="<%= rsGetFreeTier5("ProductID") %>" id="tier-5-<%=index%>">
								<img class="img-fluid" src="https://bafthumbs-400.bodyartforms.com/<%= rsGetFreeTier5("picture") %>" alt="<%=(rsGetFreeTier5("title"))%>" title="<%=(rsGetFreeTier5("title"))%>">
								<div class="w-100 text-center px-1 variation-text">
									<div class="small">
											<%= rsGetFreeTier5("title") %>
									</div>
									<div class="small font-weight-bold  d-block">
											<%= rsGetFreeTier5("min_gauge") %>
											<% if rsGetFreeTier5("min_gauge") <> rsGetFreeTier5("max_gauge") then %> 
											- <%= rsGetFreeTier5("max_gauge") %>
											<% end if %>
									</div>
								</div> 
								<div class="w-100 text-center px-1 selected-variation-text" style="display:none">
								</div>
							</a>
						<% rsGetFreeTier5.MoveNext()
						Loop
						%>
					</div>		
				</div>
			  </div>
			</div>
		  </div>
		  <!-- Accordion card -->		

		  
<%
' ------- Get FREE items for TIER 6
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "SELECT DISTINCT FlatProducts.min_gauge, FlatProducts.max_gauge, jewelry.title, jewelry.picture,jewelry.ProductID, jewelry.picture, CASE WHEN jewelry.ProductID = 2890 THEN '1' ELSE jewelry.ProductID END " & _
							"FROM ProductDetails INNER JOIN jewelry ON ProductDetails.ProductID = jewelry.ProductID " & _ 
							"INNER JOIN FlatProducts ON FlatProducts.ProductID = jewelry.ProductID " & _ 
							"WHERE ProductDetails.free <= 100 AND (jewelry.ProductID <> 3704) AND (ProductDetails.qty > 0) AND (ProductDetails.free <> 0) AND (ProductDetails.free IS NOT NULL) AND (ProductDetails.free_item_expiration_date > GETDATE() OR ProductDetails.free_item_expiration_date is Null) AND (ProductDetails.free IS NOT NULL) AND (ProductDetails.active = 1) " & _
							"ORDER BY CASE WHEN jewelry.ProductID = 2890 THEN '1' ELSE jewelry.ProductID END, jewelry.title"
		Set rsGetFreeTier6 = objCmd.Execute()
' ------- End getting free items for TIER 6
%>
		  
		  <!-- Accordion card -->
		  <div class="card" style="overflow: visible;">

			<!-- Card header -->
			<div class="card-header" role="tab" id="headingThree6">
			  <a data-toggle="collapse" data-parent="#accordionEx" href="#collapseThree6"
				aria-expanded="false" aria-controls="collapseThree6" class="accordion-button collapsed">
				<div class="d-inline-block">
				<h5 class="mb-0">
				 <%If free_items_count  < 6 Then %><i class="fa fa-lock"></i><%End If%> $100.00 CART VALUE <i class="fas fa-angle-down rotate-icon"></i>
				</h5>
				</div>
				<div id="header-card-selected-item-6"></div>				
			  </a>
			</div>

			<!-- Card body -->
			<div id="collapseThree6" class="collapse" data-tier="6" role="tabpanel" aria-labelledby="headingThree6" data-parent="#accordionEx">
			  <div class="card-body">
				<div class="mt-2 full-width-block">
					<div class="select-product-variation" id="select-product-variation-6" style="z-index: 2000; position: absolute; width: 80%; top: 117px; left: 0; right: 0; margin: auto; padding:0 !important" class="modal-body"></div>
					<div class="baf-carousel slick-free-items <%If free_items_count < 6 Then %>notavailable<%End If%>" id="slick-free-items-6" style="">
						<% rsGetFreeTier6.MoveFirst %>	
						<% index = 0 %>
						<% Do While NOT rsGetFreeTier6.EOF %>	
							<% index = index + 1 %>
							<a class="slide text-dark homepage-graphic" data-tier="6" data-slide-index="<%=index%>" data-productid="<%= rsGetFreeTier6("ProductID") %>" id="tier-6-<%=index%>">
								<img class="img-fluid" src="https://bafthumbs-400.bodyartforms.com/<%= rsGetFreeTier6("picture") %>" alt="<%=(rsGetFreeTier6("title"))%>" title="<%=(rsGetFreeTier6("title"))%>">
								<div class="w-100 text-center px-1 variation-text">
									<div class="small">
											<%= rsGetFreeTier6("title") %>
									</div>
									<div class="small font-weight-bold  d-block">
											<%= rsGetFreeTier6("min_gauge") %>
											<% if rsGetFreeTier6("min_gauge") <> rsGetFreeTier6("max_gauge") then %> 
											- <%= rsGetFreeTier6("max_gauge") %>
											<% end if %>
									</div>
								</div> 
								<div class="w-100 text-center px-1 selected-variation-text" style="display:none">
								</div>
							</a>
						<% rsGetFreeTier6.MoveNext()
						Loop
						%>
					</div>		
				</div>
			  </div>
			</div>
		  </div>
		  <!-- Accordion card -->	
		  

<%
' ------- Get FREE items for TIER 7
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "SELECT DISTINCT FlatProducts.min_gauge, FlatProducts.max_gauge, jewelry.title, jewelry.picture,jewelry.ProductID, jewelry.picture, CASE WHEN jewelry.ProductID = 2890 THEN '1' ELSE jewelry.ProductID END " & _
							"FROM ProductDetails INNER JOIN jewelry ON ProductDetails.ProductID = jewelry.ProductID " & _ 
							"INNER JOIN FlatProducts ON FlatProducts.ProductID = jewelry.ProductID " & _ 
							"WHERE ProductDetails.free <= 150 AND (jewelry.ProductID <> 3704) AND (ProductDetails.qty > 0) AND (ProductDetails.free <> 0) AND (ProductDetails.free IS NOT NULL) AND (ProductDetails.free_item_expiration_date > GETDATE() OR ProductDetails.free_item_expiration_date is Null) AND (ProductDetails.free IS NOT NULL) AND (ProductDetails.active = 1) " & _
							"ORDER BY CASE WHEN jewelry.ProductID = 2890 THEN '1' ELSE jewelry.ProductID END, jewelry.title"
		Set rsGetFreeTier7 = objCmd.Execute()
' ------- End getting free items for TIER 7
%>
		  
		  <!-- Accordion card -->
		  <div class="card" style="overflow: visible;">

			<!-- Card header -->
			<div class="card-header" role="tab" id="headingThree7">
			  <a data-toggle="collapse" data-parent="#accordionEx" href="#collapseThree7"
				aria-expanded="false" aria-controls="collapseThree7" class="accordion-button collapsed">
				<div class="d-inline-block">
				<h5 class="mb-0">
				  <%If free_items_count  < 7 Then %><i class="fa fa-lock"></i><%End If%> $150.00 CART VALUE <i class="fas fa-angle-down rotate-icon"></i>
				</h5>
				</div>
				<div id="header-card-selected-item-7"></div>				
			  </a>
			</div>

			<!-- Card body -->
			<div id="collapseThree7" class="collapse" data-tier="7" role="tabpanel" aria-labelledby="headingThree7" data-parent="#accordionEx">
			  <div class="card-body">
				<div class="mt-2 full-width-block">
					<div class="select-product-variation" id="select-product-variation-7" style="z-index: 2000; position: absolute; width: 80%; top: 117px; left: 0; right: 0; margin: auto; padding:0 !important" class="modal-body"></div>
					<div class="baf-carousel slick-free-items <%If free_items_count < 7 Then %>notavailable<%End If%>" id="slick-free-items-7" style="">
						<% rsGetFreeTier7.MoveFirst %>	
						<% index = 0 %>
						<% Do While NOT rsGetFreeTier7.EOF %>	
							<% index = index + 1 %>
							<a class="slide text-dark homepage-graphic" data-tier="7" data-slide-index="<%=index%>" data-productid="<%= rsGetFreeTier7("ProductID") %>" id="tier-7-<%=index%>">
								<img class="img-fluid" src="https://bafthumbs-400.bodyartforms.com/<%= rsGetFreeTier7("picture") %>" alt="<%=(rsGetFreeTier7("title"))%>" title="<%=(rsGetFreeTier7("title"))%>">
								<div class="w-100 text-center px-1 variation-text">
									<div class="small">
											<%= rsGetFreeTier7("title") %>
									</div>
									<div class="small font-weight-bold  d-block">
											<%= rsGetFreeTier7("min_gauge") %>
											<% if rsGetFreeTier7("min_gauge") <> rsGetFreeTier7("max_gauge") then %> 
											- <%= rsGetFreeTier7("max_gauge") %>
											<% end if %>
									</div>
								</div> 
								<div class="w-100 text-center px-1 selected-variation-text" style="display:none">
								</div>
							</a>
						<% rsGetFreeTier7.MoveNext()
						Loop
						%>
					</div>		
				</div>
			  </div>
			</div>

		  </div>
		  <!-- Accordion card -->	

		</div>
		<!-- Accordion wrapper -->	
<%
Set rsGetFreeTier1 = Nothing
Set rsGetFreeTier2 = Nothing
Set rsGetFreeTier3 = Nothing
Set rsGetFreeTier4 = Nothing
Set rsGetFreeTier5 = Nothing
%>			