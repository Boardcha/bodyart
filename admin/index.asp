<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
' set cookie to show live/sandbox mode message only for admin users
Response.Cookies("adminuser") = "yes"
Response.Cookies("adminuser").Path = "/"
Response.Cookies("adminuser").Expires =  DATE + 300
				
%>


<html>
<head>
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Administration</title>
</head>
<body>
<!--#include file="admin_header.asp"-->
<div class="p-2">

	
<div class="container-fluid p-0 mt-4">
	<div class="row no-gutters">
		<div class="col-3 px-2 pb-3">
            <div class="card h-100">
                <div class="card-header">
                  <h5>Invoice Search</h5>
                </div>
                <div class="card-body">
                    <form class="form-inline" action="/admin/invoice.asp" method="post">
						<input class="form-control form-control" type="text" name="invoice_num" placeholder="Inovice #">
						<button class="btn btn-sm btn-secondary ml-2" type="submit">Search</button>
					</form>
					<form class="form-inline" action="/admin/order history.asp" method="get">
						<input class="form-control form-control" type="text" name="var_email" placeholder="E-mail">
						<button class="btn btn-sm btn-secondary ml-2" type="submit">Search</button>
					</form>	
					<form class="form-inline" name="/admin/website_search" method="get" action="order history.asp">
						<input class="form-control form-control mr-2" name="var_first" type="text" placeholder="First name" size="10">
						<input class="form-control form-control" name="var_last" type="text" placeholder="Last name" size="10" ><br>
						<button class="btn btn-sm btn-secondary mt-1" type="submit">Search</button>
					</form>	
				
                </div><!-- card body -->
              </div><!-- card -->
        </div>
	  <div class="col-3 pr-2 pb-3">
		<div class="card h-100">
			<div class="card-header">
			  <h5>Product Search</h5>
			</div>
			<div class="card-body">
				<form class="form-inline" name="invoice_search" action="/admin/product-edit.asp" method="get">
					<input class="form-control form-control" name="ProductID" type="text" placeholder="Product ID #">
					<button class="btn btn-sm btn-secondary ml-2" type="submit">Search</button>
				</form>
				<form class="form-inline" name="detailid_search" action="/admin/SearchDetailID.asp" method="post">
					<input class="form-control form-control" name="DetailID" type="text" placeholder="Detail ID #">
					<button class="btn btn-sm btn-secondary ml-2" type="submit">Search</button>
				</form>
				<form class="form-inline" name="location_search" action="/admin/location_search.asp" method="post">
					<select class="form-control mb-1" name="section" id="search-section">
						<% While NOT rs_getsections.EOF %>                          
						<option value="<%=(rs_getsections.Fields.Item("ID_Description").Value)%>"><%=(rs_getsections.Fields.Item("ID_Description").Value)%></option>
					  <% 
					  rs_getsections.MoveNext()
					  Wend
					  %> 
				  </select>
					<input class="form-control form-control" name="location" type="text" placeholder="Location #">
					<button class="btn btn-sm btn-secondary ml-2" type="submit">Search</button>
				</form>
			</div><!-- card body -->
		  </div><!-- card -->
	</div><!-- column -->
	<div class="col-3 pr-2 pb-3">
		<div class="card h-100">
			<div class="card-header">
			  <h5>Quick Links</h5>
			</div>
			<div class="card-body">
				<a href="https://docs.google.com/spreadsheets/d/1JVi8sfV5FaEPgVNvwkAFPTr-DoJxJuBT2iebQL9FBOE/" target="_blank"><img class="mr-1" src="/images/icons/google-drive.png" height="20px">Office schedule</a><br/>
				<a href="https://docs.google.com/document/d/18lAn0wGZpdA8ufFIaIvod2mwTChk96qYuZ_FYPYwRzA/" target="_blank"><img class="mr-1" src="/images/icons/google-drive.png" height="20px">Bodyartforms handbook</a><br/>
			</div><!-- card body -->
		  </div><!-- card -->
	</div><!-- column-->
	</div><!-- row -->
  </div><!-- container -->




</div><!-- padding -->
</body>
</html>
<%
rsGetUser.Close()
Set rsGetUser = Nothing
%>