<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/template/inc_includes_ajax.asp" -->
<%

set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "UPDATE TBLWaitingList SET waiting_qty = ? WHERE ID = ?"
objCmd.Parameters.Append(objCmd.CreateParameter("waiting_qty",3,1,10, request.form("waiting_qty") ))
objCmd.Parameters.Append(objCmd.CreateParameter("WaitingID",3,1,10, request.form("waiting_id") ))
objCmd.Execute()


DataConn.Close()
Set DataConn = Nothing
%>
