<%@LANGUAGE="VBSCRIPT" %>
<!--#include virtual="/Connections/bodyartforms_sql_ADMIN.asp" -->

<%
   if request.form("department") = "cs" then     
        mailer_type = "notify-cs"

        set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "UPDATE TBLReviews SET cs_flagged = 1 WHERE ReviewID = ?"
		objCmd.Parameters.Append(objCmd.CreateParameter("ReviewID",3,1,20,Request.Form("reviewid")))
    objCmd.Execute()
    
    end if 

    if request.form("department") = "photography" then     
    mailer_type = "notify-photography"
end if 
        %>
  <!--#include virtual="/emails/function-send-email.asp"-->
  <!--#include virtual="/emails/email_variables.asp"-->
<%
DataConn.Close()
%>