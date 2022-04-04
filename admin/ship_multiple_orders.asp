<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
temp = Replace( Request.Form("Checkbox"), "'", "''" ) 
varID = Split( temp, ", " ) 

For i = 0 To UBound(varID) 

set rsGetRecords = Server.CreateObject("ADODB.Recordset")
rsGetRecords.ActiveConnection = MM_bodyartforms_sql_STRING
rsGetRecords.Source = "SELECT ID, preorder, anodize FROM sent_items WHERE ID = " & varID(i)
rsGetRecords.CursorLocation = 3 'adUseClient
rsGetRecords.LockType = 1 'Read-only records
rsGetRecords.Open()
rsGetRecords_numRows = 0

set commUpdate = Server.CreateObject("ADODB.Command")
commUpdate.ActiveConnection = MM_bodyartforms_sql_STRING
commUpdate.CommandText = "UPDATE sent_items SET shipped = 'Pending shipment' WHERE ID = " & varID(i) 
commUpdate.Execute()

'========= ALWAYS SET CUSTOM ANODIZATION TO IN PROGRESS ====================
if rsGetRecords("anodize") = true  then
      
      set commUpdate = Server.CreateObject("ADODB.Command")
      commUpdate.ActiveConnection = MM_bodyartforms_sql_STRING
      objCmd.CommandText = "UPDATE sent_items SET shipped = 'CUSTOM COLOR IN PROGRESS' WHERE ID = ?"
      objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,20, varID(i) ))
      objCmd.Execute()

end if

'========= ALWAYS PUSH CUSTOM ORDERS TO BE REVIEWED NO MATTER WHAT ====================
if rsGetRecords("preorder") = 1 then

   set commUpdate = Server.CreateObject("ADODB.Command")
   commUpdate.ActiveConnection = MM_bodyartforms_sql_STRING
   commUpdate.CommandText = "UPDATE sent_items SET shipped = 'CUSTOM ORDER IN REVIEW' WHERE ID = " & varID(i) 
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
