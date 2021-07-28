<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
'==== PAGE HAS BEEN BOOSTRAPPED =======
bootstrapped = "yes"


if request.querystring("companySort") = "" then
companySort = " "
else
companySort = "AND brandname = '" + Server.HTMLEncode(Request.Querystring("companySort")) + "'"
end if

if request.querystring("FilterMaterial") = "" then
FilterMaterial = " "
else
FilterMaterial = "AND material LIKE '%" + Server.HTMLEncode(Request.Querystring("FilterMaterial")) + "%'"
end if
%>
<%
Dim rsGetCompany
Dim rsGetCompany_numRows

Set rsGetCompany = Server.CreateObject("ADODB.Recordset")
rsGetCompany.ActiveConnection = MM_bodyartforms_sql_STRING
rsGetCompany.Source = "SELECT companyID, name, display_Header FROM dbo.TBL_Companies WHERE display_Header = 'yes' AND type = 'jewelry' ORDER BY name ASC"
rsGetCompany.CursorLocation = 3 'adUseClient
rsGetCompany.LockType = 1 'Read-only records
rsGetCompany.Open()

rsGetCompany_numRows = 0
%>
<%
Dim rsGetActive
Dim rsGetActive_numRows

Set rsGetActive = Server.CreateObject("ADODB.Recordset")
rsGetActive.ActiveConnection = MM_bodyartforms_sql_STRING
rsGetActive.Source = "SELECT ProductID, jewelry, type, title, active, brandname, material  FROM dbo.jewelry  WHERE jewelry = '" + Request.QueryString("jewelry") + "' AND active = 1 AND customorder <> 'yes' " + companySort + " " + FilterMaterial + " ORDER BY title ASC"
rsGetActive.CursorLocation = 3 'adUseClient
rsGetActive.LockType = 1 'Read-only records
rsGetActive.Open()

rsGetActive_numRows = 0
%>
<%
Dim rsGetInactive
Dim rsGetInactive_numRows

Set rsGetInactive = Server.CreateObject("ADODB.Recordset")
rsGetInactive.ActiveConnection = MM_bodyartforms_sql_STRING
rsGetInactive.Source = "SELECT ProductID, jewelry, type, title, active,brandname, material  FROM dbo.jewelry  WHERE jewelry = '" + Request.QueryString("jewelry") + "' AND active = 0 " + companySort + " " + FilterMaterial + " ORDER BY title ASC"
rsGetInactive.CursorLocation = 3 'adUseClient
rsGetInactive.LockType = 1 'Read-only records
rsGetInactive.Open()

rsGetInactive_numRows = 0
%>
<%
Dim rsGetPreorder
Dim rsGetPreorder_numRows

Set rsGetPreorder = Server.CreateObject("ADODB.Recordset")
rsGetPreorder.ActiveConnection = MM_bodyartforms_sql_STRING
rsGetPreorder.Source = "SELECT ProductID, jewelry, type, title, active, brandname, material  FROM dbo.jewelry  WHERE jewelry = '" + Request.QueryString("jewelry") + "' AND active = 1 AND customorder = 'yes' " + companySort + " " + FilterMaterial + " ORDER BY title ASC"
rsGetPreorder.CursorLocation = 3 'adUseClient
rsGetPreorder.LockType = 1 'Read-only records
rsGetPreorder.Open()

rsGetPreorder_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
rsGetActive_numRows = rsGetActive_numRows + Repeat1__numRows
%>
<%
Dim Repeat2__numRows
Dim Repeat2__index

Repeat2__numRows = -1
Repeat2__index = 0
rsGetPreorder_numRows = rsGetPreorder_numRows + Repeat2__numRows
%>
<%
Dim Repeat3__numRows
Dim Repeat3__index

Repeat3__numRows = -1
Repeat3__index = 0
rsGetInactive_numRows = rsGetInactive_numRows + Repeat3__numRows

' RESET COMPANY DROP DOWN
rsGetCompany2_numRows = 0
%>
<%
Dim Repeat4__numRows
Dim Repeat4__index

Repeat4__numRows = -1
Repeat4__index = 0
rsGetCompany_numRows = rsGetCompany_numRows + Repeat4__numRows
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Product listing</title>
</head>
<body>
<!--#include file="admin_header.asp"-->

<div class="p-3">

<table class="table table-sm table-hover table-striped">
  <thead class="thead-dark">
    <tr>
      <th colspan="3">
        ACTIVE <%= request.querystring("jewelry") %>
      </th>
    </tr>
  </thead>
  <% 
While ((Repeat1__numRows <> 0) AND (NOT rsGetActive.EOF)) 
%>
<% ' Get gauges for products
Dim rsGetGauge
Dim rsGetGauge_cmd
Dim rsGetGauge_numRows

Set rsGetGauge_cmd = Server.CreateObject ("ADODB.Command")
rsGetGauge_cmd.ActiveConnection = MM_bodyartforms_sql_STRING
rsGetGauge_cmd.CommandText = "SELECT * FROM dbo.QRY_ProductGauge WHERE ProductID = " & rsGetActive.Fields.Item("ProductID").Value & " ORDER BY GaugeOrder ASC" 
rsGetGauge_cmd.Prepared = true

Set rsGetGauge = rsGetGauge_cmd.Execute
%>
<% If Not rsGetGauge.EOF Or Not rsGetGauge.BOF Then %>
<% MinGaugeDisplay = rsGetGauge.Fields.Item("Gauge").Value %>
<% end if %>
        <%
Set rsGetGaugeMax_cmd = Server.CreateObject ("ADODB.Command")
rsGetGaugeMax_cmd.ActiveConnection = MM_bodyartforms_sql_STRING
rsGetGaugeMax_cmd.CommandText = "SELECT * FROM dbo.QRY_ProductGauge WHERE ProductID = " & rsGetActive.Fields.Item("ProductID").Value & " ORDER BY GaugeOrder DESC" 
rsGetGaugeMax_cmd.Prepared = true

Set rsGetGaugeMax = rsGetGaugeMax_cmd.Execute
%>
<% If Not rsGetGaugeMax.EOF Or Not rsGetGaugeMax.BOF Then %>
<% MaxGaugeDisplay = rsGetGaugeMax.Fields.Item("Gauge").Value %>
<% end if %>
        <%
rsGetGaugeMax.Close()
Set rsGetGaugeMax = Nothing
%>
        <%
rsGetGauge.Close()
Set rsGetGauge = Nothing
%>
    <tr>
      <td width="50%"><a href="product-edit.asp?ProductID=<%=(rsGetActive.Fields.Item("ProductID").Value)%>&info=less" class="ContentLinks"><%=(rsGetActive.Fields.Item("title").Value)%></a></td>
      <td width="25%"><%=(rsGetActive.Fields.Item("brandname").Value)%></td>
      <td width="25%"><%= MinGaugeDisplay %>
<% if MinGaugeDisplay <> MaxGaugeDisplay then %> thru <%= MaxGaugeDisplay %>
<% end if %></td>
    </tr>
    <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsGetActive.MoveNext()
Wend
%>
</table>


<table class="table table-sm table-hover table-striped">
  <thead class="thead-dark">
    <tr>
      <th colspan="3">
        PRE-ORDERS <%= request.querystring("jewelry") %>
      </th>
    </tr>
  </thead>
  <% 
While ((Repeat2__numRows <> 0) AND (NOT rsGetPreorder.EOF)) 
%>
<% ' Get gauges for products
Set rsGetGauge_cmd = Server.CreateObject ("ADODB.Command")
rsGetGauge_cmd.ActiveConnection = MM_bodyartforms_sql_STRING
rsGetGauge_cmd.CommandText = "SELECT * FROM dbo.QRY_ProductGauge WHERE ProductID = " & rsGetPreorder.Fields.Item("ProductID").Value & " ORDER BY GaugeOrder ASC" 
rsGetGauge_cmd.Prepared = true

Set rsGetGauge = rsGetGauge_cmd.Execute
%>
<% If Not rsGetGauge.EOF Or Not rsGetGauge.BOF Then %>
<% MinGaugeDisplay = rsGetGauge.Fields.Item("Gauge").Value %>
<% end if %>
        <%
Set rsGetGaugeMax_cmd = Server.CreateObject ("ADODB.Command")
rsGetGaugeMax_cmd.ActiveConnection = MM_bodyartforms_sql_STRING
rsGetGaugeMax_cmd.CommandText = "SELECT * FROM dbo.QRY_ProductGauge WHERE ProductID = " & rsGetPreorder.Fields.Item("ProductID").Value & " ORDER BY GaugeOrder DESC" 
rsGetGaugeMax_cmd.Prepared = true

Set rsGetGaugeMax = rsGetGaugeMax_cmd.Execute
%>
<% If Not rsGetGaugeMax.EOF Or Not rsGetGaugeMax.BOF Then %>
<% MaxGaugeDisplay = rsGetGaugeMax.Fields.Item("Gauge").Value %>
<% end if %>
        <%
rsGetGaugeMax.Close()
Set rsGetGaugeMax = Nothing
%>
        <%
rsGetGauge.Close()
Set rsGetGauge = Nothing
%>
    <tr>
      <td width="50%"><a href="product-edit.asp?ProductID=<%=(rsGetPreorder.Fields.Item("ProductID").Value)%>&info=less" class="ContentLinks"><%=(rsGetPreorder.Fields.Item("title").Value)%></a></td>
      <td width="25%"><%=(rsGetPreorder.Fields.Item("brandname").Value)%></td>
      <td width="25%"><%= MinGaugeDisplay %>
        <% if MinGaugeDisplay <> MaxGaugeDisplay then %>
thru <%= MaxGaugeDisplay %>
<% end if %></td>
    </tr>
    <% 
  Repeat2__index=Repeat2__index+1
  Repeat2__numRows=Repeat2__numRows-1
  rsGetPreorder.MoveNext()
Wend
%>
</table>


<table class="table table-sm table-hover table-striped">
  <thead class="thead-dark">
    <tr>
      <th colspan="3">
        IN-ACTIVE <%= request.querystring("jewelry") %>
      </th>
    </tr>
  </thead>
  <% 
While ((Repeat3__numRows <> 0) AND (NOT rsGetInactive.EOF)) 
%>
<% ' Get gauges for products
Set rsGetGauge_cmd = Server.CreateObject ("ADODB.Command")
rsGetGauge_cmd.ActiveConnection = MM_bodyartforms_sql_STRING
rsGetGauge_cmd.CommandText = "SELECT * FROM dbo.QRY_ProductGauge WHERE ProductID = " & rsGetInactive.Fields.Item("ProductID").Value & " ORDER BY GaugeOrder ASC" 
rsGetGauge_cmd.Prepared = true

Set rsGetGauge = rsGetGauge_cmd.Execute
%>
<% If Not rsGetGauge.EOF Or Not rsGetGauge.BOF Then %>
<% MinGaugeDisplay = rsGetGauge.Fields.Item("Gauge").Value %>
<% end if %>
        <%
Set rsGetGaugeMax_cmd = Server.CreateObject ("ADODB.Command")
rsGetGaugeMax_cmd.ActiveConnection = MM_bodyartforms_sql_STRING
rsGetGaugeMax_cmd.CommandText = "SELECT * FROM dbo.QRY_ProductGauge WHERE ProductID = " & rsGetInactive.Fields.Item("ProductID").Value & " ORDER BY GaugeOrder DESC" 
rsGetGaugeMax_cmd.Prepared = true

Set rsGetGaugeMax = rsGetGaugeMax_cmd.Execute
%>
<% If Not rsGetGaugeMax.EOF Or Not rsGetGaugeMax.BOF Then %>
<% MaxGaugeDisplay = rsGetGaugeMax.Fields.Item("Gauge").Value %>
<% end if %>
        <%
rsGetGaugeMax.Close()
Set rsGetGaugeMax = Nothing
%>
        <%
rsGetGauge.Close()
Set rsGetGauge = Nothing
%>
    <tr>
      <td width="50%"><a href="product-edit.asp?ProductID=<%=(rsGetInactive.Fields.Item("ProductID").Value)%>&info=less" class="ContentLinks"><%=(rsGetInactive.Fields.Item("title").Value)%></a></td>
      <td width="25%"><%=(rsGetInactive.Fields.Item("brandname").Value)%></td>
      <td width="25%"><%= MinGaugeDisplay %>
        <% if MinGaugeDisplay <> MaxGaugeDisplay then %>
thru <%= MaxGaugeDisplay %>
<% end if %></td>
    </tr>
    <% 
  Repeat3__index=Repeat3__index+1
  Repeat3__numRows=Repeat3__numRows-1
  rsGetInactive.MoveNext()
Wend
%>
</table>


</div>
</body>
</html>
<%
rsGetActive.Close()
Set rsGetActive = Nothing
%>
<%
rsGetInactive.Close()
Set rsGetInactive = Nothing
%>
<%
rsGetPreorder.Close()
Set rsGetPreorder = Nothing
%>
