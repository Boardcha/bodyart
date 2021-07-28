<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/Connections/bodyartforms_sql_ADMIN.asp" -->
<style>
 /* style sheet for "letter" printing */
    @page {
        size: 1.2in .85in;
        margin: 0
    }

    body {font-family:Arial, Helvetica, sans-serif;margin:0}

    .label {
        width:  1.2in;
        height: .5in;
        overflow: hidden;
    }

    .font-large {
        font-size: 9px;
        font-weight:bold;
    }

    .font-small {
        font-size: 7px;
    }
    
</style>
<%
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = MM_bodyartforms_sql_STRING
objCmd.CommandText = "SELECT TOP (5) * FROM tbl_po_details where po_orderid = 6157"
Set rsGetLabels = objCmd.Execute()

While NOT rsGetLabels.EOF
%>
<html>
    <head>
        <meta charset="UTF-8">
    </head>
    <body>
        <svg class="barcode"
        jsbarcode-format="CODE128"
        jsbarcode-height="20"
        jsbarcode-width="1"
        jsbarcode-margin="0"
        jsbarcode-displayValue="false"
        jsbarcode-value="<%= rsGetLabels.Fields.Item("po_orderid").Value %>.<%= rsGetLabels.Fields.Item("po_detailid").Value %>"
      </svg>
<div class="label" style="page-break-after: always">
    <span class="font-large"><%= rsGetLabels.Fields.Item("po_orderid").Value %>.<%= rsGetLabels.Fields.Item("po_detailid").Value %><br>
    QTY: CASE 0<br></span>
    <span class="font-small">Small text here, testing text wrapping, etc, etc etc etc large font threading words</span>
</div>
<%
rsGetLabels.MoveNext()
Wend
%>
</body>
<script type="text/javascript" src="../scripts/JsBarcode.all.min.js"></script>
<script charset="UTF-8">
    JsBarcode(".barcode").init();
</script>
</html>
