<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<% response.Buffer=false
Server.ScriptTimeout=300 %>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
'==== PAGE HAS BEEN BOOSTRAPPED =======
bootstrapped = "yes"

var_po_id = request("po_id")

set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn  
objCmd.CommandText = "SELECT ProductDetails.ProductDetailID, jewelry.title, ProductDetails.ProductDetail1, jewelry.ProductID, ProductDetails.Gauge, ProductDetails.Length, jewelry.picture, ProductDetails.location, TBL_Barcodes_SortOrder.ID_Description, ProductDetails.BinNumber_Detail, tbl_po_details.po_detailid, jewelry.brandname, tbl_edits_log.description FROM ProductDetails INNER JOIN jewelry ON ProductDetails.ProductID = jewelry.ProductID INNER JOIN TBL_Barcodes_SortOrder ON ProductDetails.DetailCode = TBL_Barcodes_SortOrder.ID_Number INNER JOIN tbl_po_details ON ProductDetails.ProductDetailID = tbl_po_details.po_detailid LEFT OUTER JOIN tbl_edits_log ON tbl_po_details.po_detailid = tbl_edits_log.detail_id WHERE tbl_edits_log.po_detailid = ?"
objCmd.Parameters.Append(objCmd.CreateParameter("po_orderid",3,1,20, var_po_id  ))
Set rsGetRestockItems = objCmd.Execute()	  
Set objCmd = Nothing
%>
<html>
<head>
<title>Process order</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body>
<!--#include file="admin_header.asp"-->
<div class="p-3">
<h5>
 Verify QC & Restocking of order BRAND NAME
 &nbsp;&nbsp;| &nbsp;&nbsp;Purchase order #<%= var_po_id %>
</h5>

<table  class="table table-sm table-striped table-hover mt-2">
	<thead class="thead-dark">  
	<tr>
            <th>Restocked on floor</th>
            <th>Location</th>
            <th class="Description">Description</th>
		  </tr>
		</thead>	
<% While NOT rsGetRestockItems.EOF %>
 <tr>
        <td>
            <% if instr(rsGetRestockItems("description"), "match") > 0 then %>
            <%= rsGetRestockItems("description") %>
            <% end if %>
        </td>
        <td>
            <%=(rsGetRestockItems.Fields.Item("location").Value)%> - <%=(rsGetRestockItems.Fields.Item("ID_Description").Value)%>
            <% if rsGetRestockItems.Fields.Item("BinNumber_Detail").Value <> 0 then %>
			 - BIN <%=(rsGetRestockItems.Fields.Item("BinNumber_Detail").Value)%>
			<% end if %>
        </td>
        <td>
            <a href="product-edit.asp?ProductID=<%=(rsGetRestockItems.Fields.Item("ProductID").Value)%>&info=less" target="_blank">
                <img src="https://bafthumbs-400.bodyartforms.com/<%=rsGetRestockItems("picture")%>" class="rounded float-left mr-2" style="height:50px;width:50px">
            <%=(rsGetRestockItems.Fields.Item("title").Value)%>&nbsp;<%=(rsGetRestockItems.Fields.Item("gauge").Value)%>&nbsp;<%=(rsGetRestockItems.Fields.Item("length").Value)%><%=(rsGetRestockItems.Fields.Item("ProductDetail1").Value)%></a>
        </td>
  </tr>               
  <% 
  rsGetRestockItems.MoveNext()
Wend
%>
 </table>

</div><!--admin content-->
</body>
</html>
<%
rsGetRestockItems.Close()
Set rsGetRestockItems = Nothing

DataConn.Close()
Set DataConn = Nothing
%>
