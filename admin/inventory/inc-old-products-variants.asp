<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
var_productid = request.form("productid")

    set objcmd = Server.CreateObject("ADODB.command")
    objcmd.ActiveConnection = DataConn
    objcmd.CommandText = "SELECT pd.ProductID, pd.Gauge, pd.Length, pd.ProductDetail1, pd.DateLastPurchased FROM ProductDetails pd inner join jewelry j ON j.ProductID = pd.ProductID where j.ProductID = ?"
    objcmd.Parameters.Append(objCmd.CreateParameter("ProductID",3,1,15, var_productid ))
    Set rsVariants = objcmd.Execute()
%>
<%= rsVariants.Fields.Item("Gauge").Value %>&nbsp;<%= rsVariants.Fields.Item("Length").Value %>&nbsp;<%= rsVariants.Fields.Item("ProductDetail1").Value %> - <%= rsVariants.Fields.Item("DateLastPurchased").Value %>