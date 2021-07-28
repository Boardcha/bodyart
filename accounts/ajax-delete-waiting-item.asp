<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/template/inc_includes_ajax.asp" -->
<%

set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "DELETE FROM TBLWaitingList WHERE ID = ?"
objCmd.Parameters.Append(objCmd.CreateParameter("WaitingID",3,1,10, request.form("waiting_id") ))
objCmd.Execute()


DataConn.Close()
Set DataConn = Nothing
%>
