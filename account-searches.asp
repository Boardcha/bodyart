<% @LANGUAGE="VBSCRIPT" %>
<%
	page_title = "Your account - Saved searches"
	page_description = "Your Bodyartforms account. Edit your profile, view orders, and more."
	page_keywords = ""
	
%>
<!--#include virtual="/functions/security.inc" -->
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
	
' Get saved searches
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "SELECT * FROM tbl_customer_searches  WHERE customer_ID = ?"
		objCmd.Parameters.Append(objCmd.CreateParameter("CustID_Cookie",3,1,10,CustID_Cookie))
		Set rsSearches = objCmd.Execute()
	
		
End if ' Only run this info if a match was found

%>


<div class="display-5">
		Saved Searches &amp; Filters
	</div>


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


<%
while not rsSearches.eof

	var_url = replace(rsSearches.Fields.Item("search_url").Value, ",", "")
	item_array = split(var_url, "&")
		
	%>
		<div class="card card-light mb-4 url_id_<%= rsSearches.Fields.Item("id").Value %>">
			<% if rsSearches.Fields.Item("search_nickname").Value <> "" then %>
			<div class="card-header p-2">
					<h6 class="p-0 m-0" id="nickname-text-<%= rsSearches.Fields.Item("id").Value %>"><%= rsSearches.Fields.Item("search_nickname").Value %></h6>
			</div>
			<% end if %>
			<div class="card-body p-2">
				<div class="mb-3 form-inline create-nickname nickname-id-<%= rsSearches.Fields.Item("id").Value %>" style="display:none" data-id="<%= rsSearches.Fields.Item("id").Value %>"><input class="form-control form-control-sm w-auto" type="text" name="nickname" id="nickname-<%= rsSearches.Fields.Item("id").Value %>" placeholder="Type nickname here"> <span class="btn btn-sm btn-purple ml-2 btn-create-nickname" data-id="<%= rsSearches.Fields.Item("id").Value %>"><% if rsSearches.Fields.Item("search_nickname").Value <> "" then %>Edit<%else%>Create<% end if %></span></div>
				<span class="btn btn-sm btn-purple btn-nickname" data-id="<%= rsSearches.Fields.Item("id").Value %>"><% if rsSearches.Fields.Item("search_nickname").Value <> "" then %>Edit<%else%>Create<% end if %> nickname</span>  <span class="btn btn-sm btn-purple delete-search" data-id="<%= rsSearches.Fields.Item("id").Value %>">Delete</span><span class="btn btn-sm btn-danger confirm-delete delete-<%= rsSearches.Fields.Item("id").Value %>" style="display:none" data-id="<%= rsSearches.Fields.Item("id").Value %>">Confirm delete</span>
				<div class="msg-id-<%= rsSearches.Fields.Item("id").Value %>" style="display:none"></div>
				<div class="d-block mt-3">
				<a class="btn btn-sm btn-outline-secondary mr-2" href="/products.asp?<%= var_url %>">
				<i class="fa fa-chevron-right  mr-2"></i>Go to search</a>
	<%	
	
		' Detect whether exclude
		exclude_notice = ""
		if instr(rsSearches.Fields.Item("search_url").Value, "exclude-material") then
			exclude_notice = " (excluded)"
		end if
	
		For j = 0 to Ubound(item_array)
		category = left(item_array(j),instr(item_array(j),"=")-1)
		' Make first letter uppercase
		category = UCase(Left(category,1)) & LCase(Right(category, Len(category) - 1))
		category = replace(category, "Jewelry", "Category")
		category = replace(category, "Flare_type", "Flares")
		category = replace(category, "Exclude-material", "Exclude materials")

		value = replace(Mid(item_array(j),instr(item_array(j),"=")+1), "+"," ")
		if value <> "" then
			value = UCase(Left(value,1)) & LCase(Right(value, Len(value) - 1))
			value = replace(value, "On", "Yes")
			value = replace(value, "%2f", "/")
			value = replace(value, "%22", """")
		end if 'value <> ""
	%>
			<%= category %>: <%=  value %><% if category = "Material" then %><%= exclude_notice %><% end if %><% if j <> Ubound(item_array) then %>,<% end if %><span class="mr-1"></span>
		
	<%
		next
		%></div><!-- building of search query-->
		</div><!-- card  body-->
	</div><!-- card -->
		<%
rsSearches.movenext
wend	
%>



<%
end if   'rsGetUser.EOF
%>  

<!--#include virtual="/bootstrap-template/footer.asp" -->
<script type="text/javascript">

	// Update delete button to confirm button
	$(".delete-search").click(function () {
		var id = $(this).attr("data-id");
		$(this).hide();
		$(".delete-" + id).show();
	});

	// Confirm delete
	$(".confirm-delete").click(function () {
		var id = $(this).attr("data-id");
		$(this).hide();
		
		$.ajax({
		method: "post",
		dataType: "json",
		url: "accounts/ajax-delete-search-url.asp",
		data: {id: id}
		})
		.done(function(json, msg) {
			$(".msg-id-" + id).html('<div class="alert alert-success">Success</div>').show();
			$(".url_id_" + id).delay(3000).fadeOut('slow');
		})
		.fail(function(json, msg) {
			$(".msg-id-" + id).html('<div class="alert alert-danger">Delete failed</div>').show();
		})
		
	});
	
	// Click create nickname button to show input field
	// Update delete button to confirm button
	$(".btn-nickname").click(function () {
		var id = $(this).attr("data-id");
		$(".nickname-id-" + id).show();
	});	
	
	// Confirm nickname creation
	$(".btn-create-nickname").click(function () {
		var id = $(this).attr("data-id");
		var nickname = $("#nickname-" + id).val();
		
		$(".nickname-id-" + id).hide();
		
		$.ajax({
		method: "post",
		dataType: "json",
		url: "accounts/ajax-save-search-nickname.asp",
		data: {id: id, nickname: nickname}
		})
		.done(function(json, msg) {
			$(".msg-id-" + id).html('<div class="alert alert-success">Success</div>').show();
			$(".msg-id-" + id).delay(4000).fadeOut('slow');
			$(".delete-search, .btn-nickname").show();
			$("#nickname-text-" + id).html(nickname);
		})
		.fail(function(json, msg) {
			$(".msg-id-" + id).html('<div class="alert alert-danger">Delete failed</div>').show();
			$(".delete-search, .btn-nickname").show();
		})
		
	});	


</script>