<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
Dim rsGetRecords
Dim rsGetRecords_numRows

Set rsGetRecords = Server.CreateObject("ADODB.Recordset")
rsGetRecords.ActiveConnection = MM_bodyartforms_sql_STRING
rsGetRecords.Source = "SELECT * FROM jewelry WHERE (customorder <> N'yes') AND (picture <> 'nopic.gif') AND (active = 1) AND (jewelry <> N'save') AND (date_added <= '" & date()+21 & "') AND (date_added > '" & date()-15 & "')  ORDER BY date_added DESC"
rsGetRecords.CursorLocation = 3 'adUseClient
rsGetRecords.LockType = 1 'Read-only records
rsGetRecords.Open()

rsGetRecords_numRows = 0
%>
<%
Dim HLooper1__numRows
HLooper1__numRows = -5
Dim HLooper1__index
HLooper1__index = 0
rsGetRecords_numRows = rsGetRecords_numRows + HLooper1__numRows
%>
<html>
<body>
<table cellpadding="5">
  <%
startrw = 0
endrw = HLooper1__index
numberColumns = 5
numrows = -1
while((numrows <> 0) AND (Not rsGetRecords.EOF))
	startrw = endrw + 1
	endrw = endrw + numberColumns
 %>
  <tr align="center" valign="top">
    <%
While ((startrw <= endrw) AND (Not rsGetRecords.EOF))
%>
<%
Dim rsGetPrice
Dim rsGetPrice_cmd
Dim rsGetPrice_numRows

Set rsGetPrice_cmd = Server.CreateObject ("ADODB.Command")
rsGetPrice_cmd.ActiveConnection = MM_bodyartforms_sql_STRING
rsGetPrice_cmd.CommandText = "SELECT * FROM dbo.QRY_ProductPrice WHERE ProductID = " & rsGetRecords.Fields.Item("ProductID").Value & "" 
rsGetPrice_cmd.Prepared = true
rsGetPrice_cmd.Parameters.Append rsGetPrice_cmd.CreateParameter("param1", 5, 1, -1, rsGetPrice__MMColParam) ' adDouble

Set rsGetPrice = rsGetPrice_cmd.Execute
rsGetPrice_numRows = 0
%> 
<% If Not rsGetPrice.EOF Or Not rsGetPrice.BOF Then %>
   <td width="100"><a href="http://www.bodyartforms.com/productdetails.asp?ProductID=<%= rsGetRecords.Fields.Item("ProductID").Value %>"><img src='http://bodyartforms-products.bodyartforms.com/<%=(rsGetRecords.Fields.Item("picture").Value)%>' width="90" height="90" border="0"></a><br>
        <font size="1" face="Verdana"><strong><%=(rsGetRecords.Fields.Item("title").Value)%></strong><br>


  <% if rsGetPrice.Fields.Item("MinPrice").Value = rsGetPrice.Fields.Item("MaxPrice").Value then %>
  <%= FormatCurrency(rsGetPrice.Fields.Item("MaxPrice").Value,2)%>
  <% else %>
  <%= FormatCurrency(rsGetPrice.Fields.Item("MinPrice").Value,2)%> thru <%= FormatCurrency(rsGetPrice.Fields.Item("MaxPrice").Value,2)%>
  <% end if %>
  <% if (rsGetPrice.Fields.Item("pair").Value) = "yes" then %>
  /pair
  <% end if %>
</span><br>
 <br>
        &nbsp;</font> </td>
		<%
rsGetPrice.Close()
Set rsGetPrice = Nothing
%>
    <%
	startrw = startrw + 1
	End If ' end Not rsGetPrice.EOF Or NOT rsGetPrice.BOF
	
	rsGetRecords.MoveNext()
	Wend
	%>
  </tr>
  <%
 numrows=numrows-1
 Wend
 %>
</table>
</p>
</body>
</html>
<%
rsGetRecords.Close()
Set rsGetRecords = Nothing
%>
