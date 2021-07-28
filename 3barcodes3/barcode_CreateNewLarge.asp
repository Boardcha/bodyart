<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Scan &amp; set large item</title>
<link href="../includes/nav.css" rel="stylesheet" type="text/css" />
</head>

<body onload="document.FRM_ItemScan.Item.focus();"  class="materialText">
<%
If request.form("Item") <> "" then

set SetLarge = Server.CreateObject("ADODB.Command")
SetLarge.ActiveConnection = MM_bodyartforms_sql_STRING
SetLarge.CommandText = "UPDATE ProductDetails SET DetailCode = 1 WHERE ProductDetailID = " + request.form("Item") + "" 
SetLarge.Execute()

end if
%>
<form id="FRM_ItemScan" name="FRM_ItemScan" method="post" action="barcode_SetLarge.asp">
  <p><strong>Create new detail # and set to large section</strong><br />
  </p>
  <p>Scan barcode #:
    <input name="Item" type="text" id="Item" size="10" />
  </p>
</form>
</body>
</html>
