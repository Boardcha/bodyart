<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
'====== Check to see which of the 2 reviewed by columns have info and fill in the first ======
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT reviewed_by_1, reviewed_by_2, ProductID FROM jewelry WHERE ProductID = ?"
objCmd.Parameters.Append(objCmd.CreateParameter("productid",3,1,10, request.form("productid")))
Set rsGetReviewed = objCmd.Execute()

if ISNULL(rsGetReviewed("reviewed_by_1")) then
    response.write "<br>UPDATE slot 1"
    sql_build = " reviewed_by_1 = '" & user_name & "', review_date_1 = '" & now() & "' "
end if
if ISNULL(rsGetReviewed("reviewed_by_2")) AND NOT ISNULL(rsGetReviewed("reviewed_by_1")) then
response.write "<br>UPDATE slot 2"
    sql_build = " reviewed_by_2 = '" & user_name & "', review_date_2 = '" & now() & "' "
end if

'====== ONLY UPDATE IF EITHER FIELD IS EMPTY =================
if ISNULL(rsGetReviewed("reviewed_by_1")) OR ISNULL(rsGetReviewed("reviewed_by_2")) then

response.write "<br>UPDATE DB"
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "UPDATE jewelry SET " & sql_build & " WHERE ProductID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("productid",3,1,10, request.form("productid")))
	objCmd.Execute()
end if
%>