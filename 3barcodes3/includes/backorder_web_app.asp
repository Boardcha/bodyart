<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<!DOCTYPE HTML>
<html>
<body>
<%
set commUpdate = Server.CreateObject("ADODB.Command")
commUpdate.ActiveConnection = MM_bodyartforms_sql_STRING
commUpdate.CommandText = "UPDATE TBL_OrderSummary SET BackorderReview = 'Y', notes = '" & request.form("item_notes") & "' WHERE OrderDetailID = " + Request.Form("bo_id") + "" 
commUpdate.Execute RecordsAffected, , adExecuteNoRecords

If err.number > 0 or RecordsAffected = 0 then
    Response.Write "Backorder <strong>FAILED</strong>"

else
	response.write "Backorder succcessful"
	Session("TimesScanned" & Request.Form("ses_id")) = 50
end if

DataConn.Close()
%>
</body>
</html>