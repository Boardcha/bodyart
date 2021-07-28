<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
' Get the total for the current purchase order
Set objCmd = Server.CreateObject ("ADODB.Command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT SUM(p_od.po_qty * d.wlsl_price) AS po_total FROM tbl_po_details AS p_od INNER JOIN ProductDetails AS d ON p_od.po_detailid = d.ProductDetailID WHERE (p_od.po_temp_id = ?)"
objCmd.Parameters.Append(objCmd.CreateParameter("po_temp_id",3,1,10,request.form("tempid")))
Set rsGetPOTotal = objCmd.Execute()



if not rsGetPOTotal.eof then
	if rsGetPOTotal.Fields.Item("po_total").Value <> "" then
	po_total = FormatNumber(rsGetPOTotal.Fields.Item("po_total").Value, -1, -2, -2, -2)
	else
		po_total = 0
	end if
end if
%>
{  
	"po_total":"<%= po_total %>",
	"tempid":"<%= request.form("tempid") %>"
}
<%
DataConn.Close()
%>