<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
set rsAllOrders = Server.CreateObject("ADODB.Recordset")
rsAllOrders.ActiveConnection = MM_bodyartforms_sql_STRING
rsAllOrders.Source = "SELECT ship_code, shipped, ID FROM sent_items WHERE ship_code = 'paid' AND (shipped = 'Pending shipment')"
rsAllOrders.CursorLocation = 3 'adUseClient
rsAllOrders.LockType = 1 'Read-only records
rsAllOrders.Open()

set rsAutoclaves = Server.CreateObject("ADODB.Recordset")
rsAutoclaves.ActiveConnection = MM_bodyartforms_sql_STRING
rsAutoclaves.Source = "SELECT ship_code, shipped, ID FROM sent_items WHERE ship_code = 'paid' AND autoclave = 1 AND shipped = 'Pending shipment'"
rsAutoclaves.CursorLocation = 3 'adUseClient
rsAutoclaves.LockType = 1 'Read-only records
rsAutoclaves.Open()

set rsGetUSPS = Server.CreateObject("ADODB.Recordset")
rsGetUSPS.ActiveConnection = MM_bodyartforms_sql_STRING
rsGetUSPS.Source = "SELECT ship_code, shipped, ID FROM sent_items  WHERE ship_code = 'paid' AND (shipping_type LIKE '%express%' OR shipping_type LIKE '%priority%') AND (shipping_type NOT LIKE '%DHL%') AND shipped = 'Pending shipment'"
rsGetUSPS.CursorLocation = 3 'adUseClient
rsGetUSPS.LockType = 1 'Read-only records
rsGetUSPS.Open()

set rsGetDHL = Server.CreateObject("ADODB.Recordset")
rsGetDHL.ActiveConnection = MM_bodyartforms_sql_STRING
rsGetDHL.Source = "SELECT ship_code, shipped, ID FROM sent_items  WHERE ship_code = 'paid' AND (shipping_type LIKE '%DHL%' OR shipping_type LIKE '%first class%') AND shipped = 'Pending shipment'"
rsGetDHL.CursorLocation = 3 'adUseClient
rsGetDHL.LockType = 1 'Read-only records
rsGetDHL.Open()


  ' Get information from checkboxes
temp = Replace( Request.Form("Employee"), "'", "''" ) 
Employee = Split( temp, ", " ) 

howMany = UBound(Employee)

'===== ASSIGN ALL ORDERS OUT TO PACKERS TO MAKE DOUBLE SURE NO ORDERS ARE MISSED =====================
For e = 0 To UBound(Employee)
    While NOT rsAllOrders.EOF
        ' Resets variable to loop again
        if e > UBound(Employee) then
            e = 0
        end if

		  set UpdateAll = Server.CreateObject("ADODB.Command")
		  UpdateAll.ActiveConnection = MM_bodyartforms_sql_STRING
		  UpdateAll.CommandText = "UPDATE sent_items SET PackagedBy = '" & Employee(e) & "' WHERE ID = " & rsAllOrders.Fields.Item("ID").Value & "" 
		  UpdateAll.Execute()

    e = e + 1
	rsAllOrders.MoveNext()
Wend
next


'===== ASSIGN ALL DHL ORDERS TO PACKERS EVENLY  ==================================
For e = 0 To UBound(Employee)
    While NOT rsGetDHL.EOF
        ' Resets variable to loop again
        if e > UBound(Employee) then
            e = 0
        end if

            set UpdateAll = Server.CreateObject("ADODB.Command")
            UpdateAll.ActiveConnection = MM_bodyartforms_sql_STRING
            UpdateAll.CommandText = "UPDATE sent_items SET PackagedBy = '" & Employee(e) & "' WHERE ID = " & rsGetDHL.Fields.Item("ID").Value & "" 
            UpdateAll.Execute()
    
        e = e + 1
        rsGetDHL.MoveNext()
    Wend
next

'===== ASSIGN ALL USPS ORDERS TO PACKERS EVENLY  ==================================
For e = 0 To UBound(Employee)
    While NOT rsGetUSPS.EOF
        ' Resets variable to loop again
        if e > UBound(Employee) then
            e = 0
        end if

            set UpdateAll = Server.CreateObject("ADODB.Command")
            UpdateAll.ActiveConnection = MM_bodyartforms_sql_STRING
            UpdateAll.CommandText = "UPDATE sent_items SET PackagedBy = '" & Employee(e) & "' WHERE ID = " & rsGetUSPS.Fields.Item("ID").Value & "" 
            UpdateAll.Execute()
    
        e = e + 1
        rsGetUSPS.MoveNext()
    Wend
next

'===== ASSIGN ALL AUTOCLAVE ORDERS TO PACKERS EVENLY. DO THIS ONE LAST SINCE IT DOESN'T MATTER WHAT SHIPPING METHOD IS CHOSEN. AUTOCLAVES OVERRIDE ANY SHIPPING METHOD FOR HOW LONG THEY TAKE TO PACKAGE  ==================================
For e = 0 To UBound(Employee)
    While NOT rsAutoclaves.EOF
        ' Resets variable to loop again
        if e > UBound(Employee) then
            e = 0
        end if

            set UpdateAll = Server.CreateObject("ADODB.Command")
            UpdateAll.ActiveConnection = MM_bodyartforms_sql_STRING
            UpdateAll.CommandText = "UPDATE sent_items SET PackagedBy = '" & Employee(e) & "' WHERE ID = " & rsAutoclaves.Fields.Item("ID").Value & "" 
            UpdateAll.Execute()
    
        e = e + 1
        rsAutoclaves.MoveNext()
    Wend
next

mailer_type = "split-orders"
%>
  <!--#include virtual="/emails/function-send-email.asp"-->
<%

Set objCmd = Server.CreateObject ("ADODB.Command")
objCmd.ActiveConnection = MM_bodyartforms_sql_STRING
objCmd.CommandText = "SELECT TOP (100) PERCENT PackagedBy, (SELECT COUNT(shipping_type) AS Expr0 FROM dbo.sent_items WHERE (PackagedBy = i.PackagedBy) AND (shipped = 'Pending shipment')) AS 'Total', (SELECT COUNT(shipping_type) AS Expr1 FROM dbo.sent_items WHERE (PackagedBy = i.PackagedBy) AND ((shipping_type LIKE '%DHL%') AND (shipped = 'Pending shipment'))) AS 'DHL', (SELECT COUNT(shipping_type) AS Expr2 FROM dbo.sent_items WHERE (PackagedBy = i.PackagedBy) AND (shipping_type LIKE '%express%' OR shipping_type LIKE '%priority%') AND (shipping_type NOT LIKE '%DHL%') AND (shipped = 'Pending shipment')) AS 'priority',  (SELECT COUNT(shipping_type) AS Expr6 FROM dbo.sent_items WHERE (PackagedBy = i.PackagedBy) AND (shipping_type like '%ups%') AND (shipped = 'Pending shipment')) AS 'UPS', (SELECT COUNT(autoclave) AS Expr7 FROM dbo.sent_items WHERE (PackagedBy = i.PackagedBy) AND autoclave = 1 AND (shipped = 'Pending shipment')) AS 'Autoclaves' FROM dbo.sent_items AS i WHERE (shipped = 'Pending shipment') GROUP BY PackagedBy ORDER BY PackagedBy"
Set rsEmailStats = objCmd.Execute

while not rsEmailStats.eof


if not rsEmailStats.eof then
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = MM_bodyartforms_sql_STRING
	objCmd.CommandText = "SELECT * FROM QRY_FaultyOrders WHERE (date_sent >= GETDATE() -30 AND date_sent <= GETDATE() ) AND (PackagedBy = ? OR pulled_by = ?) ORDER BY item_problem ASC, date_sent DESC"
    objCmd.Parameters.Append(objCmd.CreateParameter("PackagedBy",200,1,30, rsEmailStats("PackagedBy") ))
    objCmd.Parameters.Append(objCmd.CreateParameter("pulled_by",200,1,30, rsEmailStats("PackagedBy") ))
	Set rsGetErrors_Details = objCmd.Execute()
end if

if rsEmailStats.Fields.Item("PackagedBy").Value <> "" AND Not IsNull(rsEmailStats.Fields.Item("PackagedBy").Value) then

'------------- GET PACKER NAME AND EMAIL  -----------------
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = MM_bodyartforms_sql_STRING
objCmd.CommandText = "SELECT name, email FROM TBL_AdminUsers WHERE name = ?"
objCmd.Parameters.Append(objCmd.CreateParameter("packer_name",200,1,30, rsEmailStats("PackagedBy") ))
Set rsGetPacker = objCmd.Execute()

var_packer = rsEmailStats("PackagedBy")
%>

    <!--#include virtual="admin/packing/inc-error-formula.asp" -->
    Date 1 <%= var_date1 %>, date 2 <%= var_date2 %>, var_packer = <%= var_packer %>
  <!--#include virtual="/emails/email_variables.asp"-->
<% 
end if

rsEmailStats.movenext()
wend


rsEmailStats.Close()
Set rsEmailStats = Nothing

%>