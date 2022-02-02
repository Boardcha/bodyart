<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
        '========= DELETE A ROW FROM THE WAITING LIST TABLE ====================================

        set objCmd = Server.CreateObject("ADODB.command")
        objCmd.ActiveConnection = DataConn  
        objCmd.CommandText = "DELETE FROM TBLWaitingList WHERE ID = ?"
        objCmd.Parameters.Append(objCmd.CreateParameter("DetailID",3,1,20, request.form("id")  ))
        objCmd.Execute()
        Set objCmd = Nothing
%>