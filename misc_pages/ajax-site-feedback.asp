<%@LANGUAGE="VBSCRIPT"%>
<%
if request.form("feedback-comments") <> "" then 

mailer_type = "website-feedback"
%>
<!--#include virtual="/emails/function-send-email.asp"-->
<!--#include virtual="/emails/email_variables.asp"-->
<%

end if ' if comments are there
%>

{ "status":"success" }