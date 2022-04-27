<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/template/inc_includes_ajax.asp" -->
<%
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT *, ISNULL(Gauge, '') + ' ' + ISNULL(Length, '') + ' ' + ISNULL(ProductDetail1, '') as OptionTitle, location, ID_Number  FROM ProductDetails INNER JOIN TBL_GaugeOrder as G ON ISNULL(ProductDetails.Gauge,'') = ISNULL(G.GaugeShow,'') INNER JOIN TBL_Barcodes_SortOrder ON ProductDetails.DetailCode = TBL_Barcodes_SortOrder.ID_Number WHERE ProductID = ? AND active = 1 ORDER BY G.GaugeOrder ASC, item_order ASC"
objCmd.Parameters.Append(objCmd.CreateParameter("ProductID",3,1,10,request("productid")))
Set rsGetItems = objCmd.Execute()
%>
<table class="table table-striped table-borderless table-hover">
    <thead class="thead-dark">
        <tr>
            <th>Location</th>
            <th>Item</th>
        </tr>
    </thead>
<%
while not rsGetItems.eof
%>
<tr>
    <td>
        <%
         if rsGetItems.Fields.Item("BinNumber_Detail").Value = 0 then %>
            <%= rsGetItems.Fields.Item("ID_Description").Value %>
        <%
            BinType = ""
        else 
             If  (rsGetItems.Fields.Item("ID_Description").Value = "Case 1" OR rsGetItems.Fields.Item("ID_Description").Value = "Case 2" OR rsGetItems.Fields.Item("ID_Description").Value = "Case 3" OR rsGetItems.Fields.Item("ID_Description").Value = "Case 4") Then  
                BinType = rsGetItems("ID_Description") & " Shelf "
            else
                BinType = "BIN "
            end if
        %>
            <%= BinType %>
                <%= rsGetItems.Fields.Item("BinNumber_Detail").Value %> 
        <% end if %>
        <% If rsGetItems.Fields.Item("BinNumber_Detail").Value <> 0 then
        '==== Show detail id for items in limited bins
        %>   
            &nbsp;&nbsp;<%= rsGetItems.Fields.Item("ProductDetailID").Value %>
        <% else 
        '===== regular stock item location
        %>
        &nbsp;&nbsp;<%= rsGetItems.Fields.Item("location").Value %>
        
        <% end if %>
    </td>
    <td>
        <%= rsGetItems("OptionTitle") %>
    </td>
</tr>
<%
rsGetItems.movenext()
wend
%>
</table>