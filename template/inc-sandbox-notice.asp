<%
' For my personal testing locally, always have it in sandbox mode so I don't fuck anythiing up
'session("sandbox") = "ON"

if request.cookies("adminuser") = "yes" then
	if request.querystring("inactive") = "yes" then
		session("inactive") = "yes"
	end if
	if request.querystring("inactive") = "no" then
		session("inactive") = ""
	end if
end if

 if request.cookies("adminuser") = "yes" and session("sandbox") <> "ON" then 
%>
	<div class="live-notice">
		LIVE MODE <span id="sandbox-toggle-on" class="sandbox-toggle">Sandbox</span>
		<span style="font-size: .8em; font-weight: normal">
		<% if session("inactive") = "" then %>
		&nbsp;&nbsp;&nbsp;<a href="?<%= Request.ServerVariables("QUERY_STRING") %>&amp;inactive=yes">Show inactives</a>
		<% end if %>
		<% if session("inactive") = "yes" then %>
		&nbsp;&nbsp;&nbsp;<a href="?<%= Request.ServerVariables("QUERY_STRING") %>&amp;inactive=no">Don't show inactives</a>
		<% end if %>
		</span>
	</div>

	<script type="text/javascript">

	// Toggle sandbox front end load
	$('#sandbox-toggle-on').click(function(){
		 $.ajax({
		  url: "sandbox.asp",
		  data: { sandbox:"ON" }
		}).done(function() {
		  location.reload();
		});		
	});

	</script>
	
<% end if %>
<% if request.cookies("adminuser") = "yes" and session("sandbox") = "ON" then %>
	<div class="sandbox-notice">
		SANDBOX TESTING MODE <span id="sandbox-toggle-off" class="sandbox-toggle">Turn off</span>
		<span style="font-size: .8em; font-weight: normal">
		<% if session("inactive") = "" then %>
		&nbsp;&nbsp;&nbsp;<a href="?<%= Request.ServerVariables("QUERY_STRING") %>&amp;inactive=yes">Show inactives</a>
		<% end if %>
		<% if session("inactive") = "yes" then %>
		&nbsp;&nbsp;&nbsp;<a href="?<%= Request.ServerVariables("QUERY_STRING") %>&amp;inactive=no">Don't show inactives</a>
		<% end if %></span>
		
	</div>
	
	<script type="text/javascript">

	// Toggle sandbox front end load
	$('#sandbox-toggle-off').click(function(){
		 $.ajax({
		  url: "sandbox.asp",
		  data: { sandbox:"OFF" }
		}).done(function() {
		  location.reload();
		});		
	});

	</script>
<% end if %>
