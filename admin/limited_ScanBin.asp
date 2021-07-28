<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Scan bin</title>
<link href="../includes/nav.css" rel="stylesheet" type="text/css" />
</head>

<body class="materialText" onload="document.FRM_InvoiceScan.Bin.focus();">

  <% if request.querystring("Complete") = "Yes" then %>
    <span class="HelpHeader">ALL ITEMS COMPLETED</span>

    <% end if ' show if invoice is completed%>
<form id="FRM_InvoiceScan" name="FRM_InvoiceScan" method="post" action="limited_GetItems.asp">
  <p class="productheaders">Scan limited BIN #</p>
  <input type="text" name="Bin" id="Bin" />
  <input name="ResetBin" type="hidden" id="ResetBin" value="yes" />
</form>
<br />
</body>
</html>
