<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
Dim rsWaitingReply
Dim rsWaitingReply_numRows

Set rsWaitingReply = Server.CreateObject("ADODB.Recordset")
rsWaitingReply.ActiveConnection = MM_bodyartforms_sql_STRING
rsWaitingReply.Source = "SELECT ID, shipped, customer_first, customer_last, comments FROM sent_items WHERE shipped = 'WAITING FOR CUSTOMER REPLY' ORDER BY customer_first ASC"
rsWaitingReply.CursorLocation = 3 'adUseClient
rsWaitingReply.LockType = 1 'Read-only records
rsWaitingReply.Open()

rsWaitingReply_numRows = 0
%>
<%
Dim rsClaim__MMColParam
rsClaim__MMColParam = "CLAIM IN PROGRESS"
If (Request("MM_EmptyValue") <> "") Then 
  rsClaim__MMColParam = Request("MM_EmptyValue")
End If
%>
<%
Dim rsClaim
Dim rsClaim_numRows

Set rsClaim = Server.CreateObject("ADODB.Recordset")
rsClaim.ActiveConnection = MM_bodyartforms_sql_STRING
rsClaim.Source = "SELECT ID, shipped, customer_first, customer_last, comments FROM sent_items WHERE shipped = 'FILED TRACE' ORDER BY customer_first ASC"
rsClaim.CursorLocation = 3 'adUseClient
rsClaim.LockType = 1 'Read-only records
rsClaim.Open()

rsClaim_numRows = 0
%>
<%
Dim rsAutoclave__MMColParam
rsAutoclave__MMColParam = "AUTOCLAVING ORDER"
If (Request("MM_EmptyValue") <> "") Then 
  rsAutoclave__MMColParam = Request("MM_EmptyValue")
End If
%>
<%
Dim rsAutoclave
Dim rsAutoclave_numRows

Set rsAutoclave = Server.CreateObject("ADODB.Recordset")
rsAutoclave.ActiveConnection = MM_bodyartforms_sql_STRING
rsAutoclave.Source = "SELECT ID, shipped, customer_first, customer_last, comments FROM sent_items WHERE shipped = '" + Replace(rsAutoclave__MMColParam, "'", "''") + "' ORDER BY customer_first ASC"
rsAutoclave.CursorLocation = 3 'adUseClient
rsAutoclave.LockType = 1 'Read-only records
rsAutoclave.Open()

rsAutoclave_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
rsWaitingReply_numRows = rsWaitingReply_numRows + Repeat1__numRows
%>
<%
Dim Repeat2__numRows
Dim Repeat2__index

Repeat2__numRows = -1
Repeat2__index = 0
rsPayment_numRows = rsPayment_numRows + Repeat2__numRows
%>
<%
Dim Repeat3__numRows
Dim Repeat3__index

Repeat3__numRows = -1
Repeat3__index = 0
rsChargebacks_numRows = rsChargebacks_numRows + Repeat3__numRows
%>
<%
Dim Repeat4__numRows
Dim Repeat4__index

Repeat4__numRows = -1
Repeat4__index = 0
rsNotArrived_numRows = rsNotArrived_numRows + Repeat4__numRows
%>
<%
Dim Repeat5__numRows
Dim Repeat5__index

Repeat5__numRows = -1
Repeat5__index = 0
rsClaim_numRows = rsClaim_numRows + Repeat5__numRows
%>
<%
Dim Repeat6__numRows
Dim Repeat6__index

Repeat6__numRows = -1
Repeat6__index = 0
rsAutoclave_numRows = rsAutoclave_numRows + Repeat6__numRows
%>
<html>
<head>
<link rel="stylesheet" type="text/css" href="../includes/nav.css" />
<title>Claims &amp; other issues</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"></head>
<body bgcolor="#666699" topmargin="0" text="#CCCCCC" link="#CCCCCC" vlink="#CCCCCC">

  <!--#include file="admin_header.asp"-->

<br>
<table width="100%" border="0" cellspacing="1" cellpadding="5">
  <tr bgcolor="#000000" valign="middle">
    <td align="left" valign="top" class="pricegauge"><strong>WAITING FOR CUSTOMER REPLIES</strong></td>
  </tr>
  <%If (Repeat1__numRows Mod 2) Then%>
  <% 
While ((Repeat1__numRows <> 0) AND (NOT rsWaitingReply.EOF)) 
%>
  <tr bgcolor="#4E4E74" valign="top" align="left">
    <td align="left" valign="top"><p><span class="pricegauge"><a href="invoice.asp?ID=<%=rsWaitingReply.Fields.Item("ID").Value %>" class="faqlinks"><%=(rsWaitingReply.Fields.Item("customer_first").Value)%> &nbsp;<%=(rsWaitingReply.Fields.Item("customer_last").Value)%> - Invoice #<%=(rsWaitingReply.Fields.Item("ID").Value)%> </a></span><span class="pricegauge"><br>
      <%=(rsWaitingReply.Fields.Item("comments").Value)%></span></p>      </td>
  </tr>
  <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsWaitingReply.MoveNext()
Wend
%>
  <%End If%>
</table>
&nbsp;&nbsp;<br>
<table width="100%" border="0" cellspacing="1" cellpadding="5">
  <tr bgcolor="#000000" valign="middle">
    <td align="left" valign="top" class="pricegauge"><strong><strong><a name="claim" id="claim"></a></strong>CLAIMS IN PROGRESS </strong></td>
  </tr>
  <% 
While ((Repeat5__numRows <> 0) AND (NOT rsClaim.EOF)) 
%>
  <tr bgcolor="#4E4E74" valign="top" align="left">
    <td align="left" valign="top"><p><span class="pricegauge"><a href="invoice.asp?ID=<%=rsClaim.Fields.Item("ID").Value %>" class="faqlinks"><%=(rsClaim.Fields.Item("customer_first").Value)%> &nbsp;<%=(rsClaim.Fields.Item("customer_last").Value)%> - Invoice #<%=(rsClaim.Fields.Item("ID").Value)%> </a></span><span class="pricegauge"><br>
      <%=(rsClaim.Fields.Item("comments").Value)%></span></p></td>
  </tr>
  <% 
  Repeat5__index=Repeat5__index+1
  Repeat5__numRows=Repeat5__numRows-1
  rsClaim.MoveNext()
Wend
%>


</table>
<br>
</body>
</html>
<%
rsWaitingReply.Close()
Set rsWaitingReply = Nothing
%>
<%
rsClaim.Close()
Set rsClaim = Nothing
%>
<%
rsAutoclave.Close()
Set rsAutoclave = Nothing
%>
