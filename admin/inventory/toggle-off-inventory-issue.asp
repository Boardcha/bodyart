<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/Connections/bodyartforms_sql_ADMIN.asp" -->
<%
' ====== FLAG ISSUE AS FIXED
Set objCmd = Server.CreateObject ("ADODB.Command")
objCmd.ActiveConnection = DataConn
If request.form("issue_id") <> "all" Then issue_id = " AND issue_id = " & request.form("issue_id")
objCmd.CommandText = "UPDATE tbl_product_issues SET issue_fixed = 1 WHERE issue_fixed = 0" & issue_id 
objCmd.Execute  
%>
