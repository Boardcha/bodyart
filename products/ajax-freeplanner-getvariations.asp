<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<!--#include virtual="/template/inc_includes_ajax.asp" -->
<%
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT jewelry.title, jewelry.picture, jewelry.picture_400, ProductDetails.ProductDetail1, jewelry.ProductID, ProductDetails.ProductDetailID, ProductDetails.Free_QTY, ProductDetails.Gauge, ProductDetails.Length, ProductDetails.detail_code, ISNULL(ProductDetails.gauge,'') + ' ' + ISNULL(ProductDetails.Length,'') + ' ' + ISNULL(ProductDetails.ProductDetail1,'') AS 'free_title' FROM ProductDetails INNER JOIN jewelry ON ProductDetails.ProductID = jewelry.ProductID WHERE jewelry.ProductID = ? AND ProductDetails.qty > 0 AND ProductDetails.free <> 0 AND ProductDetails.free IS NOT NULL AND ProductDetails.active = 1 ORDER BY item_order ASC"
objCmd.Parameters.Append(objCmd.CreateParameter("productid",3,1,50, request("productid") ))
Set rsGetFreeVariations = objCmd.Execute()
%>
<img class="img-fluid" src="https://bafthumbs-400.bodyartforms.com/<%= rsGetFreeVariations("picture_400") %>"  title="<%= rsGetFreeVariations("title") %>" alt="<%= rsGetFreeVariations("title") %>" />
<br>
<%
while NOT rsGetFreeVariations.eof
%>
Qty <%= rsGetFreeVariations("Free_QTY") %>&nbsp;&nbsp;--&nbsp;&nbsp;<%= rsGetFreeVariations("free_title") %><br>
<%
rsGetFreeVariations.movenext()
wend
%>
<%
DataConn.Close()
Set DataConn = Nothing
%>