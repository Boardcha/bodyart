<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/Connections/bodyartforms_sql_ADMIN.asp" -->
<%
' ====== FLAG ISSUE AS FIXED
Set objCmd = Server.CreateObject ("ADODB.Command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "UPDATE tbl_product_issues SET issue_fixed = 1 WHERE issue_id = ?"
objCmd.Parameters.Append(objCmd.CreateParameter("issue_id",3,1,15, request.form("issue_id") ))
objCmd.Execute  
%>
