<% @LANGUAGE="VBSCRIPT" %>
<!--#include virtual="/template/inc_includes.asp" -->

<% 'if Request.form("email") <> "" then 

mailer_type = "reported-photo"
				
%>

<!--#include virtual="emails/function-send-email.asp"-->
<!--#include virtual="emails/email_variables.asp"-->

<div class="alert alert-success">
	Thank you! Your report has been sent to customer service. 
	<br/>
	If you need more assistance accessing your account feel free to contact our <a class="alert-link"  href="/contact.asp">customer service department</a> and we'll be happy to help!
</div>
