<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
Dim rsGetInvoice__MMColParam
rsGetInvoice__MMColParam = "1"
if (Request.form("invoice_num") <> "") then 
rsGetInvoice__MMColParam = Request.form("invoice_num") 
else
rsGetInvoice__MMColParam = Request.QueryString("ID")
end if
%>
<%
set rsGetInvoice = Server.CreateObject("ADODB.Recordset")
rsGetInvoice.ActiveConnection = MM_bodyartforms_sql_STRING
rsGetInvoice.Source = "SELECT *  FROM sent_items  WHERE ID = " + Replace(rsGetInvoice__MMColParam, "'", "''") + ""
rsGetInvoice.CursorLocation = 3 'adUseClient
rsGetInvoice.LockType = 1 'Read-only records
rsGetInvoice.Open()
rsGetInvoice_numRows = 0
%>




<title>INVOICE</title>
<style type="text/css">
<!--
.style1 {font-size: 16px; font-weight: bold; font-family: "Century Gothic";}
-->
</style>
<link href="../includes/nav.css" rel="stylesheet" type="text/css">
<body bgcolor="#FFFFFF" text="#000000">
<table width="650" border="0" cellspacing="0" cellpadding="1" bgcolor="#000000">
  <tr>
    <td>
      <table width="100%" border="0" cellspacing="0" cellpadding="5" bgcolor="#FFFFFF">
        <tr>
          <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td width="65%" align="left" valign="top"><p><font face="Arial" size="2"><b><font face="Arial" size="2"><b><img src="../images/bodyartforms_face.gif" width="99" height="87" align="left"><img src="../images/bodyartforms_text.gif" width="214" height="35"></b></font><br>
                  </b></font><span class="materialText">1966 S. Austin Ave.<br>
                    Georgetown, TX  78626
                  </span><font face="Arial" size="2"><b><br>
                    </b></font></p></td>
              <td width="35%" align="left" valign="top"><font color="#000000"><span class="style1">INVOICE 
                # <%=(rsGetInvoice.Fields.Item("ID").Value)%></span></font><br>
                  <span class="materialText"><%=(rsGetInvoice.Fields.Item("customer_first").Value)%>&nbsp;<%=(rsGetInvoice.Fields.Item("customer_last").Value)%><br>
                        <%=(rsGetInvoice.Fields.Item("address").Value)%> <br>
                    <%=(rsGetInvoice.Fields.Item("city").Value)%>, <%=(rsGetInvoice.Fields.Item("state").Value)%>&nbsp;<%=(rsGetInvoice.Fields.Item("province").Value)%>&nbsp; <%=(rsGetInvoice.Fields.Item("zip").Value)%><br>
                      <%=(rsGetInvoice.Fields.Item("country").Value)%></span></td>
            </tr>
          </table>
            <font size="4" face="Verdana, Arial, Helvetica, sans-serif"><strong>&nbsp;<br>
            Tracking # <%=(rsGetInvoice.Fields.Item("UPS_tracking").Value)%></strong></font><br>
            &nbsp;<br>
            <table width="100%" border="0" cellspacing="0" cellpadding="2">
              <tr>
                <td width="60%" class="materialText"><strong>Item description</strong></td>
                <td width="15%" class="materialText"><strong>WHOLESALE  each</strong></td>
                <td width="15%" class="materialText"><strong>Price </strong></td>
              </tr>
              <%
Dim rsGetOrderDetails
Dim rsGetOrderDetails_numRows

Set rsGetOrderDetails = Server.CreateObject("ADODB.Recordset")
With rsGetOrderDetails
rsGetOrderDetails.ActiveConnection = MM_bodyartforms_sql_STRING
rsGetOrderDetails.Source = "SELECT OrderDetailID, qty, title, ProductDetail1, PreOrder_Desc, wlsl_price FROM dbo.QRY_OrderDetails WHERE ID = " & rsGetInvoice.Fields.Item("ID").Value & ""
rsGetOrderDetails.CursorLocation = 3 'adUseClient
rsGetOrderDetails.LockType = 1 'Read-only records
rsGetOrderDetails.Open()

LineItem = 0
SumLineItem = 0

Do While Not.Eof
%>
              <tr>
                <td width="60%"><font size="2" face="Arial"><%=(rsGetOrderDetails.Fields.Item("qty").Value)%>&nbsp; |&nbsp; <%=(rsGetOrderDetails.Fields.Item("title").Value)%>&nbsp;&nbsp;<%=(rsGetOrderDetails.Fields.Item("ProductDetail1").Value)%>
                </font></td>
                <td width="15%"><font size="2" face="Arial"><%= FormatCurrency((rsGetOrderDetails.Fields.Item("wlsl_price").Value), -1, -2, -0, -2) %></font></td>
                <td width="15%"><font size="2" face="Arial"><%= FormatCurrency((rsGetOrderDetails.Fields.Item("wlsl_price").Value)*(rsGetOrderDetails.Fields.Item("qty").Value), -1, -2, -0, -2) %></font></td>
              </tr>
              <%
LineItem = rsGetOrderDetails.Fields.Item("wlsl_price").Value * rsGetOrderDetails.Fields.Item("qty").Value
SumLineItem = SumLineItem + LineItem

.Movenext()
Loop
End With 

rsGetOrderDetails.Close()
Set rsGetOrderDetails = Nothing
rsGetOrderDetails_numRows = 0
%>
              <tr>
                <td colspan="5" align="left"><p>&nbsp;</p>
                  <p><b><font size="2" face="Century Gothic">TOTAL:&nbsp;
                    $<%= SumLineItem %> + Shipping </font></b> </p></td>
              </tr>
            </table>
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
          </table></td>
        </tr>
    </table></td>
  </tr>
</table>
<p>
<%
rsGetInvoice.Close()
%>
</p>

