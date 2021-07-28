<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
Dim rsGetDetail__MMColParam
rsGetDetail__MMColParam = "1"
If (Request.Querystring("DetailID") <> "") Then 
  rsGetDetail__MMColParam = Request.Querystring("DetailID")
End If
%>
<%
Dim rsGetDetail
Dim rsGetDetail_cmd
Dim rsGetDetail_numRows

Set rsGetDetail_cmd = Server.CreateObject ("ADODB.Command")
rsGetDetail_cmd.ActiveConnection = MM_bodyartforms_sql_STRING
rsGetDetail_cmd.CommandText = "SELECT * FROM dbo.QRY_Barcode_PullInvoice WHERE OrderDetailID = ?" 
rsGetDetail_cmd.Prepared = true
rsGetDetail_cmd.Parameters.Append rsGetDetail_cmd.CreateParameter("param1", 5, 1, -1, rsGetDetail__MMColParam) ' adDouble

Set rsGetDetail = rsGetDetail_cmd.Execute
rsGetDetail_numRows = 0
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Scan bin</title>
<link href="../../includes/nav.css" rel="stylesheet" type="text/css" />
</head>

<body class="materialText" onload="document.FRM_InvoiceScan.Bin.focus();">

  <% if request.querystring("Complete") = "Yes" then %>
    <span class="HelpHeader">ALL ITEMS COMPLETED</span>

    <% end if ' show if invoice is completed%>
<form id="FRM_InvoiceScan" name="FRM_InvoiceScan" method="post" action="limited_GetItems.asp">
  <p class="productheaders">Scan limited BIN #</p>
  <input type="text" name="Bin" id="Bin" />
</form>
<br />
</body>
</html>
