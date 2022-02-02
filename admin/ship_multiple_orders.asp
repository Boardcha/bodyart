<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
temp = Replace( Request.Form("Checkbox"), "'", "''" ) 
varID = Split( temp, ", " ) 

set commUpdate = Server.CreateObject("ADODB.Command")
commUpdate.ActiveConnection = MM_bodyartforms_sql_STRING

For i = 0 To UBound(varID) 

set rsGetRecords = Server.CreateObject("ADODB.Recordset")
rsGetRecords.ActiveConnection = MM_bodyartforms_sql_STRING
rsGetRecords.Source = "SELECT ID, preorder FROM sent_items WHERE ID = " & varID(i)
rsGetRecords.CursorLocation = 3 'adUseClient
rsGetRecords.LockType = 1 'Read-only records
rsGetRecords.Open()
rsGetRecords_numRows = 0

if (rsGetRecords.Fields.Item("preorder").Value) = 1 then

commUpdate.CommandText = "UPDATE sent_items SET shipped = 'CUSTOM ORDER IN REVIEW' WHERE ID = " & varID(i) 
    ' comment out next line AFTER IT WORKS 
    'Response.Write "DEBUG SQL: " & commUpdate.CommandText & "<BR/>" 
commUpdate.Execute()

else

commUpdate.CommandText = "UPDATE sent_items SET shipped = 'Pending shipment' WHERE ID = " & varID(i) 
    ' comment out next line AFTER IT WORKS 
    'Response.Write "DEBUG SQL: " & commUpdate.CommandText & "<BR/>" 
commUpdate.Execute()

end if

rsGetRecords.Close()
Set rsGetRecords = Nothing
   Next
   
if request.form("status") = "review" then
   Response.Redirect("review_orders.asp")
end if

if request.form("status") = "pre-order" then
   Response.Redirect("custom_orders.asp")
end if
%>
<html>
<head>
<title>Records are being updated</title>
</head>
<body bgcolor="#666699" text="#CCCCCC">
Records are being updated
</body>
</html>
