<%@ Language=VBScript %>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
'==== PAGE HAS BEEN BOOSTRAPPED =======
bootstrapped = "yes"

	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT * FROM TBL_AdminUsers WHERE archived = 0 ORDER BY name ASC"
	Set rs_getUser = objCmd.Execute()
	
%>
<html>
<head>
<!--#include file="includes/inc_scripts.asp"-->
<title>
	Manage admin users
</title>
</head>
<body>
<!--#include file="admin_header.asp"-->
<div class="p-3">
<% If var_access_level = "Admin" or var_access_level = "Manager" then 
%>
<form class="admin-fields ajax-update">
<table class="table table-borderless table-hover">
	<thead class="thead-dark">
		<tr>
			<th>Name</th>
			<th>User name</th>
			<th>Access level</th>
			<th>Email</th>
			<th style="text-align:center">Reset password</th>
			<th style="width:5%">Assign orders</th>
			<th style="text-align:center">Archive user</th>
		</tr>
	</thead>	
	<tbody class="alert alert-info">
		<tr>
			<td>
				<input class="form-control form-control-sm" type="text" id="add_name">
			</td>
			<td>
				<input class="form-control form-control-sm" type="text" id="add_username">
			</td>
			<td>
				<select class="form-control form-control-sm" id="add_access_level">
					<option value="Customer service">Customer service</option>
					<option value="Inventory">Inventory</option>
					<option value="Packaging">Packaging</option>
					<option value="Photography">Photography</option>
					<option value="Pre-orders">Pre-orders</option>
					<option value="Social Media">Social Media</option>
					<option value="Admin">Admin</option>
				</select>
			</td>
			<td style="text-align:center">
				<button class="btn btn-sm btn-info" id="btn-add-user"  type="button">Add new user</button>
			</td>
			<td></td>
			<td></td>
			<td></td>
		</tr>
	</tbody>
<%
		While NOT rs_getUser.EOF
%>
	<tbody id="<%= rs_getUser.Fields.Item("ID").Value %>">
		<tr>
			<td>
				<input class="form-control form-control-sm" type="text" name="name_<%= rs_getUser.Fields.Item("ID").Value %>" value="<%= rs_getUser.Fields.Item("name").Value %>" data-column="name" data-id="<%= rs_getUser.Fields.Item("ID").Value %>">
			</td>
			<td>
				<input class="form-control form-control-sm" type="text" name="username_<%= rs_getUser.Fields.Item("ID").Value %>" value="<%= rs_getUser.Fields.Item("username").Value %>" data-column="username" data-id="<%= rs_getUser.Fields.Item("ID").Value %>">
			</td>
			<td>
				<select class="form-control form-control-sm" name="access_level_<%= rs_getUser.Fields.Item("ID").Value %>" data-column="AccessLevel" data-id="<%= rs_getUser.Fields.Item("ID").Value %>">
					<option value="<%= rs_getUser.Fields.Item("AccessLevel").Value %>"><%= rs_getUser.Fields.Item("AccessLevel").Value %></option>
					<option value="Customer service">Customer service</option>
					<option value="Inventory">Inventory</option>
					<option value="Packaging">Packaging</option>
					<option value="Photography">Photography</option>
					<option value="Pre-orders">Pre-orders</option>
					<option value="Manager">Manager</option>
					<option value="Admin">Admin</option>
				</select>
			</td>
			<td>
				<input class="form-control form-control-sm" type="text" name="email_<%= rs_getUser.Fields.Item("ID").Value %>" value="<%= rs_getUser.Fields.Item("email").Value %>" data-column="email" data-id="<%= rs_getUser.Fields.Item("ID").Value %>">
			</td>
			<td style="text-align:center">
				<select class="form-control form-control-sm" name="password" data-id="<%= rs_getUser.Fields.Item("ID").Value %>" data-username="<%= rs_getUser.Fields.Item("name").Value %>">
					<option value="">Select recipient</option>
					<option value="amanda@bodyartforms.com" data-name="Amanda">Email to Amanda</option>
					<option value="ellen@bodyartforms.com" data-name="Ellen">Email to Ellen</option>
					<option value="andres@bodyartforms.com" data-name="Andres">Email to Andres</option>
				</select>
			</td>
			<td style="text-align:center">
				<div class="custom-control custom-checkbox">
					<input name="toggle_packer_<%= rs_getUser("ID") %>" id="toggle_packer_<%= rs_getUser("ID") %>" class="custom-control-input" type="checkbox" value="1" <% if rs_getUser("toggle_packer") = 1 then %>checked<% end if %> data-unchecked="0" data-column="toggle_packer"  data-id="<%= rs_getUser("ID") %>">
					<label class="custom-control-label" for="toggle_packer_<%= rs_getUser("ID") %>"></label>
				  </div>
			</td>
			<td style="text-align:center">
				<button class="btn btn-sm btn-outline-warning btn-archive-user" data-id="<%= rs_getUser.Fields.Item("ID").Value %>"  type="button">Archive user</button>
			</td>
		</tr>
	</tbody>
<% 
	rs_getUser.MoveNext()
	Wend
%>
</table>
</form>
<%
else ' access denied %>
Access denied
<% end if %>
</div>
</body>
</html>
<script type="text/javascript">
	$(document).ready(function(){
	
		// Send reset password link to selected user
		$("select[name='password']").change(function(){
			var column_val = $(this).val();
			var id = $(this).attr("data-id");
			var emailname = $('option:selected', this ).attr("data-name");
			var username = $(this).attr("data-username");
			
			$.ajax({
			method: "POST",
			url: "administrative/ajax_reset_user_password.asp",
			data: {id: id, email: column_val, emailname: emailname, username: username}
			})
			.done(function( msg ) {
				console.log("Success");
	
			})
			.fail(function(msg) {
				console.log("fail");
			});
		});

		// Archive user
		$(".btn-archive-user").click(function(){
			var user_id = $(this).attr("data-id");
			
			$.ajax({
			method: "POST",
			url: "administrative/ajax-archive-user.asp",
			data: {user_id: user_id}
			})
			.done(function( msg ) {
				$('#' + user_id).fadeOut('slow');	
			})
			.fail(function(msg) {
				alert("Archive failed");
			});
		});

		// Add new admin user
		$("#btn-add-user").click(function(){
			var var_name = $('#add_name').val();
			var var_username = $('#add_username').val();
			var var_access_level = $('#add_access_level').val();
			
			$.ajax({
			method: "POST",
			url: "administrative/ajax-add-user.asp",
			data: {var_name: var_name, var_username: var_username, var_access_level: var_access_level}
			})
			.done(function( msg ) {
				location.reload();
			})
			.fail(function(msg) {
				alert("Archive failed");
			});
		});
	
		function auto_update() {
			// Auto-update form fields
			$(".ajax-update select:not([name='password'], #add_access_level), .ajax-update input:not('#add_name, #add_username')").change(function(){
	
				var column_name = $(this).attr("data-column");
				var column_val = $(this).val();
				var id = $(this).attr("data-id");
				var field_name = $(this).attr("name");
				var friendly_name = $(this).attr("data-friendly");
				
				if ($(this).is(':checkbox')) {
					if ($(this).prop("checked")) { // Get values if it's a checkbox
						column_val = $(this).val();
					} else {
						column_val = $(this).attr("data-unchecked");
					}
				}
				
				var $this = $(this);
				if ($this.is("input")) {
					var field_type = "input"
				} else if ($this.is("select")) {
					var field_type = "select"
				} else if ($this.is("textarea")) {
					var field_type = "textarea"
				}
					
				$.ajax({
					method: "POST",
					url: "administrative/ajax_update_user.asp",
					data: {id: id, column: column_name, value: column_val,  friendly_name:friendly_name}
					})
					.done(function( msg ) {
						// Highlight field green for success
						$(field_type + "[name='"+ field_name +"']").addClass("alert-success");
						
						setTimeout(function(){					
							$(field_type + "[name='"+ field_name +"']").removeClass("alert-success");
						}, 3000);	
							
							
						//	console.log("Success");
					//	alert( "error" + msg + "Column: " + column_name + " Value: " + column_val + " ID: " + id + "Detail ID: " + detail_id  );
					})
					.fail(function(msg) {
					// Highlight field red for failure
					
					$(field_type + "[name='"+ field_name +"']").addClass("alert-danger");
						setTimeout(function(){
							$(field_type + "[name='"+ field_name +"']").removeClass("alert-danger");
						}, 3000);					
						alert("The field did not save. Try again or contact Amanda.");
							
					//	console.log( "error" + msg + "Column: " + column_name + " Value: " + column_val + " ID: " + id + "Detail ID: " + detail_id  );
					});
			});
		} // end auto update function
		
		auto_update(); // run function
	});
	</script>
<%
DataConn.Close()
Set DataConn = Nothing
%>
