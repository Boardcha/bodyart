<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
varItem = Request.Form("Item")

If request.form("BinUpdate") = "yes" then

If len(varItem) > 4 then
	Response.Write "<b>Error scanning into bin. Re-scan the <u>ITEM</u> and try again.</b>"
else

set comm = Server.CreateObject("ADODB.Command")
comm.ActiveConnection = MM_bodyartforms_sql_STRING
comm.CommandText = "UPDATE ProductDetails SET BinNumber_Detail = " & request.form("Item") & " WHERE ProductDetailID = " & request.form("DetailID") & "" 
comm.Execute()
Set comm = Nothing

Response.Write "Scanned into bin # " & varItem & "<br/>Scan next item"

End if

varItem = ""

end if %>
<%
Dim rsGetBinNumber__MMColParam
rsGetBinNumber__MMColParam = "1"
If (varItem <> "") Then 
  rsGetBinNumber__MMColParam = varItem
End If
%>
<%
Dim rsGetBinNumber
Dim rsGetBinNumber_cmd
Dim rsGetBinNumber_numRows

Set rsGetBinNumber_cmd = Server.CreateObject ("ADODB.Command")
rsGetBinNumber_cmd.ActiveConnection = MM_bodyartforms_sql_STRING
rsGetBinNumber_cmd.CommandText = "SELECT ProductDetailID, BinNumber_Detail FROM dbo.ProductDetails WHERE ProductDetailID = ?" 
rsGetBinNumber_cmd.Prepared = true
rsGetBinNumber_cmd.Parameters.Append rsGetBinNumber_cmd.CreateParameter("param1", 5, 1, -1, rsGetBinNumber__MMColParam) ' adDouble

Set rsGetBinNumber = rsGetBinNumber_cmd.Execute
rsGetBinNumber_numRows = 0
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Scan &amp; set to bin number</title>
<style type="text/css">
<!--
body {
		-webkit-text-size-adjust:none;
	  	font-family: Helvetica, Arial, Verdana, sans-serif;
	  	font-size: 15px;
	  	color: black;
	  }
	
.alert {
	   color: #CC0000;
	   font-weight: bold;
	   font-size: 25px;
	   }
-->
</style>
</head>

<body>
<form id="FRM_BinScan" name="FRM_BinScan" method="post" action="barcode_SetBinNumber.asp">
    <% If Not rsGetBinNumber.EOF Or Not rsGetBinNumber.BOF Then %>
<% if (rsGetBinNumber.Fields.Item("BinNumber_Detail").Value) <> 0 then %>
<span class="alert">ALREADY IN BIN # <%=(rsGetBinNumber.Fields.Item("BinNumber_Detail").Value)%></span>
<% end if %>
<% If rsGetBinNumber.Fields.Item("BinNumber_Detail").Value = "0" then %>
Scan bin #
<input name="BinUpdate" type="hidden" id="BinUpdate" value="yes" />
<input name="DetailID" type="hidden" id="DetailID" value="<%=(rsGetBinNumber.Fields.Item("ProductDetailID").Value)%>" />
<% End if %>
<% end if ' if recordset is not empty %><input name="Item" type="hidden" id="Item" size="10" />
</form>
</body>
</html>
<%
rsGetBinNumber.Close()
Set rsGetBinNumber = Nothing
%>
