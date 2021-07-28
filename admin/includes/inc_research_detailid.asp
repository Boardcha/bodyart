<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
if request.querystring("detailid") <> "" then ' check to see if a detailid is provided and if not, just update the main products table	
	
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT dbo.ProductDetails.qty, dbo.ProductDetails.DateLastPurchased, dbo.ProductDetails.BinNumber_Detail, dbo.jewelry.active, jewelry.picture, jewelry.title FROM dbo.jewelry INNER JOIN                     dbo.ProductDetails ON dbo.jewelry.ProductID = dbo.ProductDetails.ProductID WHERE (dbo.ProductDetails.ProductDetailID = ?)"
	objCmd.Parameters.Append(objCmd.CreateParameter("detailid",3,1,10,request.querystring("detailid")))
	Set rsResearch = objCmd.Execute()

if rsResearch.Fields.Item("active").Value = 1 then
	active = "Yes"
else
	active = "No"
end if

end if
%>
<p>
<img src="http://bodyartforms-products.bodyartforms.com/<%= rsResearch.Fields.Item("picture").Value %>" />
<br/><strong><%= rsResearch.Fields.Item("title").Value %></strong>
</p>
<p>
<strong>Detail ID:</strong> <%= request.querystring("detailid") %><br/>
<strong>Bin #:</strong> <%= rsResearch.Fields.Item("BinNumber_Detail").Value %><br/>
<strong>Quantity:</strong> <%= rsResearch.Fields.Item("qty").Value %><br/>
<strong>Active:</strong> <%= active %><br/>
<strong>Last sold:</strong> <%= rsResearch.Fields.Item("DateLastPurchased").Value %>
</p>

<%
DataConn.Close()
Set rsResearch = Nothing
%>