<!--#include virtual="/template/inc_includes_ajax.asp" -->
<%
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "DELETE FROM wishlist WHERE ID = ?"
objCmd.Parameters.Append(objCmd.CreateParameter("id",3,1,12, request.form("wishlist_id")))
objCmd.Execute()


DataConn.Close()
Set DataConn = Nothing
%>
