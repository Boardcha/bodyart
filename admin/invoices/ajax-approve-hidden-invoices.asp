<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
if request.form("push_hidden") = "partial" then
    var_sql = " TOP(25) "
end if

If WeekDayName(WeekDay(date())) = "Saturday" OR WeekDayName(WeekDay(date())) = "Sunday" then
    sql_delay_150s = " AND over_150 <> 1 "
end if

set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = MM_bodyartforms_sql_STRING
objCmd.CommandText = "UPDATE " & var_sql & " sent_items SET shipped = 'Pending shipment' WHERE ship_code = 'paid' AND shipped = 'Pending...' "  & sql_delay_150s & " AND (Review_OrderError <> 1 OR Review_OrderError IS NULL)"
objCmd.Execute()


' ===== Count all hidden records to be shipped
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = MM_bodyartforms_sql_STRING
objCmd.CommandText = "SELECT * FROM sent_items WHERE ship_code = 'paid' AND shipped = 'Pending...'"
set rsGetHiddenRecords = Server.CreateObject("ADODB.Recordset")
rsGetHiddenRecords.CursorLocation = 3 'adUseClient
rsGetHiddenRecords.Open objCmd
hidden_total = rsGetHiddenRecords.RecordCount
%>
{
    "records_total":"<%= hidden_total %>"
}
<%
				
DataConn.Close()
%>