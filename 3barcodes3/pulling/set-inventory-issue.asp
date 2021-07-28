<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/Connections/bodyartforms_sql_ADMIN.asp" -->
<%
' ====== UPDATE ORDER WITH ANY INVENTORY PROBLEMS THAT WERE SUBMITTED
Set objCmd = Server.CreateObject ("ADODB.Command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "UPDATE TBL_OrderSummary SET inventory_issue_toggle = 1, inventory_issue_description = ? WHERE OrderDetailID = ?"
objCmd.Parameters.Append(objCmd.CreateParameter("notes",200,1,3000, request.form("item_issue") & " " & request.form("notes") ))
objCmd.Parameters.Append(objCmd.CreateParameter("orderdetailid",3,1,15, request.form("orderdetailid") ))
objCmd.Execute  
%>
