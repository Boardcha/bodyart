<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->

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
					<!--#include virtual="/admin/landing/inc-invoice-search.inc"-->				
                </div><!-- card body -->
              </div><!-- card -->
        </div>
	  <div class="col-3 pr-2 pb-3">
		<div class="card h-100">
			<div class="card-header">
			  <h5>Product Search</h5>
			</div>
			<div class="card-body">
				<!--#include virtual="/admin/landing/inc-product-searches.inc"-->
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