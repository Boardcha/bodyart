<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
	' Get results
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT ProductID, title, picture from jewelry WHERE ProductID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("productid",3,1,10,request("productid")))
	set rsGetProduct = objCmd.Execute()	
	
	' Get results
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT * FROM ProductDetails LEFT OUTER JOIN TBL_GaugeOrder ON ProductDetails.Gauge = TBL_GaugeOrder.GaugeShow  WHERE ProductID = ? and active = 1 and qty > 0  ORDER BY TBL_GaugeOrder.GaugeOrder"
	objCmd.Parameters.Append(objCmd.CreateParameter("productid",3,1,10,request("productid")))
	set rsGetDetails = objCmd.Execute()



%>
<img src="http://bodyartforms-products.bodyartforms.com/<%=(rsGetProduct.Fields.Item("picture").Value)%>" class="thumb50"><br/>
<input type="hidden" id="frm_productid" value="<%=(rsGetProduct.Fields.Item("ProductID").Value)%>">
<div class="font-weight-bold">
		<%= rsGetProduct.Fields.Item("title").Value %>
</div>

<%
while not rsGetDetails.eof
%>
<div class="my-1">
	<span class="btn btn-sm btn-outline-secondary mr-3 select_item_to_add bo-exchange-item" data-add_detail="<%= rsGetDetails.Fields.Item("ProductDetailID").Value %>" data-add_price="<%= rsGetDetails.Fields.Item("price").Value %>" data-itemname="<%= rsGetDetails.Fields.Item("gauge").Value %>&nbsp;<%= rsGetDetails.Fields.Item("length").Value %>&nbsp;<%= rsGetDetails.Fields.Item("ProductDetail1").Value %>">
	<%= rsGetDetails.Fields.Item("gauge").Value %>&nbsp;<%= rsGetDetails.Fields.Item("length").Value %>&nbsp;<%= rsGetDetails.Fields.Item("ProductDetail1").Value %>
	</span>
	(<%= rsGetDetails.Fields.Item("qty").Value %> in stock)
</div>
<%
rsGetDetails.movenext()
wend

DataConn.Close()
%>