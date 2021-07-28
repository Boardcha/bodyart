<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/Connections/bodyartforms_sql_ADMIN.asp" -->

<%
    set objCmd = Server.CreateObject("ADODB.command")
    objCmd.ActiveConnection = DataConn
    objCmd.CommandText = "UPDATE jewelry SET pull_completed = 1 WHERE ProductID = ?"
    objCmd.Parameters.Append(objCmd.CreateParameter("productid",3,1,15, request.form("productid") ))
    set rsGetDetails = objCmd.Execute()

%>
