<%@LANGUAGE="VBSCRIPT"  CODEPAGE="65001"%>
<!--#include virtual="/Connections/sql_connection.asp" -->
<%
	
title = "Free Items"	
title_onpage = "Free Items"
description = "Free Items"
var_photo_url = "https://bafthumbs-400.bodyartforms.com"

SQL = _
"SELECT * FROM(" & _
	"SELECT jewelry.ProductID, jewelry.title, ProductDetails.ProductDetailID, ProductDetails.price, ProductDetails.qty, ProductDetails.ProductDetail1, ProductDetails.Gauge, jewelry.picture, ProductDetails.detail_code, ProductDetails.Free_QTY, ProductDetails.free, " & _
	"ROW_NUMBER() OVER (PARTITION BY jewelry.ProductID, ProductDetails.free ORDER BY jewelry.ProductID DESC) AS [ROWNUMBER], " & _
	"Count(jewelry.ProductID) OVER (PARTITION BY jewelry.ProductID) AS [ROWCOUNT] " & _
	"FROM ProductDetails INNER JOIN jewelry ON ProductDetails.ProductID = jewelry.ProductID " & _
	"WHERE (jewelry.ProductID <> 3704) AND (ProductDetails.qty > 0) AND (ProductDetails.free <> 0) AND (ProductDetails.free IS NOT NULL) AND (ProductDetails.active = 1)) GROUPS " & _
"WHERE GROUPS.[ROWNUMBER] = 1 ORDER BY GROUPS.free, GROUPS.ProductDetailID"
	  
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = SQL
Set rsGetRecords = objCmd.Execute()
%>
<!--#include virtual="/bootstrap-template/header-connection.asp" -->
<!--#include virtual="/bootstrap-template/header-scripts-and-css.asp" -->
<!--#include virtual="/bootstrap-template/header-json-schemas.asp" -->
<!--#include virtual="/bootstrap-template/header-navigation.asp" -->
<!--#include virtual="/bootstrap-template/filters.asp" -->

<div class="products">
	<div class="display-5" style="font-size:1.6em">
		<%= title_onpage %>
	</div>
	<% if NOT rsGetRecords.EOF then	%>
		<div class="d-flex flex-row flex-wrap">
			<% While NOT rsGetRecords.EOF %>
				<%If treshold <> rsGetRecords("free") Then%>
					<div class="display-5 mb-2 w-100" style="font-size:1.6em; margin-top:35px">
						<%= "$" & FormatNumber(rsGetRecords("free"), 2) & " CART THRESHOLD"%>
					</div>
				<%End If%>
				
				<% treshold = rsGetRecords("free") %>
				<div class="col-12 col-xs-4 col-md-3 col-xl-3 col-break1600-2 my-3 px-1 px-md-2 text-center">	
					<div class="container-fluid">
						<div class="row border-bottom border-secondary">	
							<a class="col p-0 text-dark" href="productdetails.asp?ProductID=<%= rsGetRecords("ProductID") %>" data-historyid="nav<%= rsGetRecords("ProductID") %>">
								<div class="position-relative">
									<img class="img-fluid w-100 <% if lazy_count > 20 then %> lazyload <% end if %>" <% if lazy_count > 20 then %> src="/images/image-placeholder.png" data-src="<%= var_photo_url %>/<%=(rsGetRecords("picture"))%>" <% else %> src="<%= var_photo_url %>/<%=(rsGetRecords("picture"))%>" <% end if %> title="<%=(rsGetRecords("title"))%>" alt="<%=(rsGetRecords("title"))%>" />
								</div>
							</a> 
						</div>
						<a class="text-dark" href="productdetails.asp?ProductID=<%= rsGetRecords("ProductID") %>" data-historyid="nav<%= rsGetRecords("ProductID") %>">
							<div class="row">
								<div class="small text-center w-100 px-1">
									<%If rsGetRecords("title") = "Order Credit" Then %>
										<%= "$" & FormatNumber(rsGetRecords("price")) & " " & rsGetRecords("title")%>
									<%Else%>
										<%=rsGetRecords("title")%>
									<%End If%>	
								</div> 
								<%If rsGetRecords("rowcount") > 1 Then %>
									<div class="small text-center w-100 px-1">
										(<%=rsGetRecords("rowcount")%> available variations)
									</div>
								<%End If%>	
							</div>
						</a>
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

<button class="products-top rounded-circle text-center position-fixed px-2 py-1 alert alert-secondary pointer" type="button"><i class="fa fa-chevron-up"></i></button>
</div><!-- end main content-box -->

<!--#include virtual="/bootstrap-template/footer.asp" -->
<script type="text/javascript" src="/js-pages/products.min.js?v=040221"></script>