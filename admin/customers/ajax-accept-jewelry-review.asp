<%@LANGUAGE="VBSCRIPT" %>
<!--#include virtual="/Connections/bodyartforms_sql_ADMIN.asp" -->

<%


	var_status = "accepted"
	' Set variable if it's anything besides accepted
	if Request.Form("vote") <> "accepted" then
		var_status = "rejected"
	end if
	
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "UPDATE dbo.TBLReviews SET comments = ?, status = ?, date_posted = ?, review = ? WHERE ReviewID = ?"
		objCmd.Prepared = true
		objCmd.Parameters.Append(objCmd.CreateParameter("Comments",200,1,500,Request.Form("comments")))
		objCmd.Parameters.Append(objCmd.CreateParameter("Vote",200,1,300,var_status))
		objCmd.Parameters.Append(objCmd.CreateParameter("DatePosted",200,1,20,date()))
		objCmd.Parameters.Append(objCmd.CreateParameter("review",200,1,4000,Request.Form("Review")))
		objCmd.Parameters.Append(objCmd.CreateParameter("ReviewID",3,1,20,Request.Form("review-id")))
		objCmd.Execute()

'IF APPROVED
if Request.Form("vote") = "accepted" then

    ' Give 1 point to customer
    set objCmd = Server.CreateObject("ADODB.Command")
    objCmd.ActiveConnection = DataConn
    objCmd.CommandText = "UPDATE customers SET Points = Points + 1 WHERE customer_ID = ?" 
    objCmd.Parameters.Append(objCmd.CreateParameter("@GetCustomerID",3,1,20,Request.Form("customer-id")))
    objCmd.Execute()

else ' IF REJECTED


		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "UPDATE dbo.TBLReviews SET status = 'rejected' WHERE ReviewID = ?"
		objCmd.Prepared = true
		objCmd.Parameters.Append(objCmd.CreateParameter("ReviewID",3,1,20,Request.Form("review-id")))
        objCmd.Execute()
        
		
		
		' only send if it's not set to no email send 
		if Request.Form("vote") <>  "rejected" then

		mailer_type = "reject-review"
        %>

        <!--#include virtual="emails/function-send-email.asp"-->
	    <!--#include virtual="emails/email_variables.asp"-->
<%
		end if ' reject reason
end if ' if <> accepted
	

Set rsGetReviews = Nothing
DataConn.Close()
%>