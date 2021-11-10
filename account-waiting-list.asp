<% @LANGUAGE="VBSCRIPT" %>
<%
	page_title = "Your account - Waitlist items"
	page_description = "Your Bodyartforms account. Edit your profile, view orders, and more."
	page_keywords = ""
	
%>
<!--#include virtual="/bootstrap-template/header-connection.asp" -->
<!--#include virtual="/bootstrap-template/header-scripts-and-css.asp" -->
<!--#include virtual="/bootstrap-template/header-json-schemas.asp" -->
<!--#include virtual="/bootstrap-template/header-navigation.asp" -->

<!--#include virtual="/bootstrap-template/filters.asp" -->
<%
var_flagged = ""

' Pull the customer information from a cookie
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "SELECT * FROM customers  WHERE customer_ID = ?"
		objCmd.Parameters.Append(objCmd.CreateParameter("CustID_Cookie",3,1,10,CustID_Cookie))
		Set rsGetUser = objCmd.Execute()
		
If Not rsGetUser.EOF Or Not rsGetUser.BOF Then ' Only run this info if a match was found

	if rsGetUser.Fields.Item("Flagged").Value = "Y" then
		var_flagged = "yes"
	end if
	
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT shipped, email FROM sent_items WHERE email = ? AND (shipped = 'Flagged' OR shipped = 'Chargeback')"
	objCmd.Parameters.Append(objCmd.CreateParameter("email",200,1,250,rsGetUser.Fields.Item("email").Value))
	set rsGetFlaggedOrders = objCmd.Execute()
	
	if NOT rsGetFlaggedOrders.eof then
		var_flagged = "yes"
	end if
	
' Get waitlist items
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "SELECT dbo.TBLWaitingList.DetailID, dbo.ProductDetails.qty, dbo.TBLWaitingList.name, dbo.TBLWaitingList.email, dbo.ProductDetails.ProductID, ISNULL(ProductDetails.Gauge, '') + ' ' + ISNULL(ProductDetails.Length, '') + ' ' + ISNULL(ProductDetails.ProductDetail1, '') + ' ' + ISNULL(jewelry.title, '') AS title, dbo.jewelry.type, dbo.TBLWaitingList.ID, jewelry.picture, jewelry.largepic, dbo.ProductDetails.ProductDetail1, dbo.TBLWaitingList.customerID, dbo.jewelry.brandname, dbo.ProductDetails.active, dbo.jewelry.active AS ActiveProduct, dbo.jewelry.customorder, dbo.ProductDetails.Gauge, dbo.ProductDetails.Length, waiting_qty, TBL_Companies.ShowTextLogo, TBL_Companies.ProductLogo FROM dbo.ProductDetails INNER JOIN dbo.TBLWaitingList ON dbo.ProductDetails.ProductDetailID = dbo.TBLWaitingList.DetailID INNER JOIN dbo.jewelry ON dbo.ProductDetails.ProductID = dbo.jewelry.ProductID INNER JOIN TBL_Companies ON jewelry.brandname = TBL_Companies.name  WHERE customerID = ? ORDER BY title ASC"
		objCmd.Parameters.Append(objCmd.CreateParameter("CustID_Cookie",3,1,10,CustID_Cookie))
		Set rsWaitingList = objCmd.Execute()
	
		
End if ' Only run this info if a match was found

%>


<div class="display-5">
		Waitlist items
	</div>
        <h6>You are currently signed up to be notified when the following items come back in stock</h6>
<%
if session("admin_tempcustid") <> "" then %>
	<div class="alert alert-success">Admin viewing
	&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	<a href="account.asp?cleartemp=yes">Reset</a>
	</div>
<% end if %>
<!--#include virtual="/accounts/inc-account-navigation.asp" -->
<% If rsGetUser.EOF or var_flagged = "yes" Then
%>
	<div class="alert alert-danger">Not logged in or no account found</div>
<% elseif rsGetUser("active") = 0 then %>
	<div class="alert alert-danger"><h5>Your account has not been activated yet.</h5>Please click on the activation link sent to your email to confirm your account registration and access your account.</div>
<% else %>


<% If Not rsWaitingList.EOF Then %>
<div class="d-flex flex-row flex-wrap">
<% Do While Not rsWaitingList.EOF %>
    <div class="col-6 col-md-4 col-lg-4 col-xl-3 col-break1600-2 my-3 px-0 px-md-2" id="block-<%= rsWaitingList.Fields.Item("ID").Value %>">	
		<div class="container-fluid p-0">
<div class="mx-1">
	<a class="mb-2 d-block" href="/productdetails.asp?ProductID=<%= rsWaitingList.Fields.Item("ProductID").Value %>">
		<div class="position-relative">
            <img class="img-fluid"  src="https://bafthumbs-400.bodyartforms.com/<%= rsWaitingList.Fields.Item("picture").Value %>"  style="width:400px;height:auto" alt="Product photo">
		
			<% if rsWaitingList.Fields.Item("ShowTextLogo").Value = "Y" then 
			%>
			<div class="brand-info position-absolute w-50 badge badge-light rounded-0" style=" bottom:5px;right: 5px;overflow-wrap: break-word;">
				<img class="img-fluid" src="/images/<%= rsWaitingList.Fields.Item("ProductLogo").Value %>" alt="logo" />
			</div>
			<% end if %>
		</div><!-- relative position-->
	</a>
	<div class="form-inline">
    <span class="btn btn-sm btn-outline-danger mr-4 delete-item" data-id="<%= rsWaitingList.Fields.Item("ID").Value %>"><i class="fa fa-trash-alt"></i></span>
    Quantity <input class="form-control form-control-sm ml-1 mr-3 update-qty" style="width: 50px" type="text" name="qty_<%= rsWaitingList.Fields.Item("ID").Value %>" value="<%= rsWaitingList.Fields.Item("waiting_qty").Value %>" data-id="<%= rsWaitingList.Fields.Item("ID").Value %>"><span class="text-success font-weight-bold" id="msg-update-<%= rsWaitingList.Fields.Item("ID").Value %>"></span>
</div>
		<a class="btn btn-purple btn-block m-0 my-1" href="/productdetails.asp?ProductID=<%=(rsWaitingList.Fields.Item("ProductID").Value)%>">View product</a>
		<div class="font-weight-bold small">
			<% If (rsWaitingList.Fields.Item("type").Value) = "Discontinued" or (rsWaitingList.Fields.Item("type").Value) = "limited" then %>									
			<span class="text-warning"><%=(rsWaitingList.Fields.Item("type").Value)%></span>&nbsp;&nbsp;
			<% end if %>
			<%=(rsWaitingList.Fields.Item("title").Value)%>
		</div>
</div><!-- margin padding -->
	</div><!-- container fluid -->
</div><!--flex columns-->



<% 
rsWaitingList.MoveNext()
Loop
End If ' end Not rsWaitingList.EOF 
%>
</div><!-- flex-row -->

<%
end if   'rsGetUser.EOF
%>  

<!--#include virtual="/bootstrap-template/footer.asp" -->
<script type="text/javascript" src="/js-pages/account-waitlist.min.js?v=031620"></script>