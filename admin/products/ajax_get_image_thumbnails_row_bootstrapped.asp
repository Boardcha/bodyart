<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%

	Set objCmd = Server.CreateObject ("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT * FROM tbl_images WHERE product_id = ? ORDER BY img_sort ASC, img_id ASC" 
	objCmd.Parameters.Append(objCmd.CreateParameter("productid",3,1,15,request.queryString("productid")))
	Set rs_GetImgID = objCmd.Execute()

If rs_GetImgID.EOF then 
%>
	<span class="font-weight-bold text-light">No additional images</span>
<%
else
	While NOT rs_GetImgID.EOF 
%>
	<img src="http://bodyartforms-products.bodyartforms.com/<%=(rs_GetImgID.Fields.Item("img_thumb").Value)%>" class="my-1 mr-1 mini-thumb thumb-activate img_<%=(rs_GetImgID.Fields.Item("img_id").Value)%>"  style="width: 30px;height: auto" id="<%= rs_GetImgID.Fields.Item("img_id").Value %>" data-imgid="<%=(rs_GetImgID.Fields.Item("img_id").Value)%>" data-name="<%=(rs_GetImgID.Fields.Item("img_full").Value)%>" data-description="<%= rs_GetImgID.Fields.Item("img_description").Value %>">
	
<%
	 rs_GetImgID.MoveNext()
	Wend
%>
<!--<span class="m-1 mini-thumb thumb-activate img_0" id="0" data-imgid="0" style="width: 30px; height: 30px; background-color: grey;display:inline-block;vertical-align:middle;"></span>-->

<%
end if
DataConn.Close()
%>