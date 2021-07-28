<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/template/inc_includes_ajax.asp" -->
<%
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT * FROM customers WHERE customer_ID = ?"
objCmd.Parameters.Append(objCmd.CreateParameter("CustID_Cookie",3,1,10,CustID_Cookie))
Set rsGetUser = objCmd.Execute()
		
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "UPDATE customers SET credits = ? + credits, Points = 0 WHERE customer_ID = ?"
objCmd.Parameters.Append(objCmd.CreateParameter("Credits",6,1,10,rsGetUser.Fields.Item("Points").Value * .05))
objCmd.Parameters.Append(objCmd.CreateParameter("CustomerID",3,1,12,CustID_Cookie))
objCmd.Execute()


' Get current store credit amount after points have been converted into it
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT credits FROM customers WHERE customer_ID = ?"
objCmd.Parameters.Append(objCmd.CreateParameter("CustID_Cookie",3,1,10,CustID_Cookie))
Set rsNewAmount = objCmd.Execute()
%>
{
	"amount":"<%= FormatCurrency(rsNewAmount.Fields.Item("credits").Value, 2) %>"
}
<%
DataConn.Close()
Set DataConn = Nothing
%>
