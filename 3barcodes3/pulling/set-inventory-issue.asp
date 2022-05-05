<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/Connections/bodyartforms_sql_ADMIN.asp" -->
<%
' ====== ADD ANY PRODUCT PROBLEM OR ISSUE TO TABLE FOR TRACKING AND RESOLVING
Set objCmd = Server.CreateObject ("ADODB.Command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "INSERT INTO tbl_product_issues (issue_detailid, issue_reported_by_who, issue_description, issue_report_source) VALUES (?,?,?,?)"
objCmd.Parameters.Append(objCmd.CreateParameter("issue_detailid",3,1,15, request.form("detailid") ))
objCmd.Parameters.Append(objCmd.CreateParameter("issue_reported_by_who",200,1,50, user_name ))
objCmd.Parameters.Append(objCmd.CreateParameter("issue_description",200,1,500, request.form("item_issue") & " " & request.form("notes") ))
objCmd.Parameters.Append(objCmd.CreateParameter("issue_report_source",200,1,50, request.form("report_source") ))
objCmd.Execute  
%>
