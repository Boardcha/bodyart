
<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/Connections/bodyartforms_sql_ADMIN.asp" -->
<%
if session("pulling_for") <> "" then
    pulling_for = session("pulling_for")
else 
    pulling_for = ""
end if

If Not rsGetUser.EOF then

Set objCmd = Server.CreateObject ("ADODB.Command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "UPDATE sent_items SET pull_order_no = 0 WHERE CAST(date_sent AS date) = CAST(GETDATE()-64 AS date)"
objCmd.Parameters.Append(objCmd.CreateParameter("puller",200,1,50, rsGetUser.Fields.Item("name").Value ))
objCmd.Execute  


    Set objCmd = Server.CreateObject ("ADODB.Command")
    objCmd.ActiveConnection = DataConn
    objCmd.CommandText = "SELECT TOP(50) ID, PackagedBy, pulled_by, pull_order_no, customer_first, customer_last, shipped FROM sent_items WHERE (sent_items.shipped = N'Shipped' or sent_items.shipped = N'Pending...' or sent_items.shipped = N'Pending shipment' or sent_items.shipped = N'Cancelled' OR sent_items.shipped = N'ON HOLD') AND pull_order_no = 0  AND (PackagedBy = ?) AND CAST(date_sent AS date) = CAST(GETDATE()-85 AS date) ORDER BY PackagedBy, CASE WHEN shipping_type LIKE '%office%' THEN 1 WHEN autoclave = 1 THEN 2 WHEN shipping_type LIKE '%express%' THEN 4 WHEN shipping_type LIKE '%ups%' THEN 5 WHEN shipping_type LIKE '%USPS Priority mail%' THEN 6 WHEN shipping_type LIKE '%max%' THEN 7 WHEN   (sent_items.shipping_type = 'DHL Basic mail') THEN 8 WHEN  (sent_items.shipping_type = 'USPS First Class Mail') THEN 9 WHEN  (sent_items.shipping_type = 'DHL GlobalMail Packet Priority') THEN 10 WHEN  (sent_items.shipping_type = 'DHL GlobalMail Parcel Priority') THEN 11 WHEN  (sent_items.shipping_type LIKE '%global basic%') THEN 12 ELSE 20 END ASC, ID ASC"
    objCmd.Parameters.Append(objCmd.CreateParameter("packer",200,1,50, pulling_for ))
    Set rsAssignInvoices = objCmd.Execute



%>

<!DOCTYPE html>
<html lang="en">
<body>
        <% if rsAssignInvoices.eof then %>
        <div class="alert alert-danger mt-3">Either no packer is selected to pull for OR no invoices are available for packer selected</div>
        <% end if %>
<table class="table tabl-sm">
<% while not rsAssignInvoices.eof %>
    <tr <% if rsAssignInvoices.Fields.Item("shipped").Value = "Cancelled" OR  rsAssignInvoices.Fields.Item("shipped").Value = "ON HOLD" then %>class="table-danger"<% end if %>>
        <td>
                <div class="custom-control custom-checkbox">
                        <% if rsAssignInvoices.Fields.Item("shipped").Value <> "Cancelled" AND  rsAssignInvoices.Fields.Item("shipped").Value <> "ON HOLD" then %>
                        <input type="checkbox" class="custom-control-input check-invoice" id="<%= rsAssignInvoices.Fields.Item("ID").Value %>" value="<%= rsAssignInvoices.Fields.Item("ID").Value %>">
                        <label class="custom-control-label" for="<%= rsAssignInvoices.Fields.Item("ID").Value %>">
                            <% end if %>
                            <span class="mr-5"><%= rsAssignInvoices.Fields.Item("ID").Value %></span>
                            <%= rsAssignInvoices.Fields.Item("customer_first").Value %>&nbsp;<%= rsAssignInvoices.Fields.Item("customer_last").Value %>
                            <% if rsAssignInvoices.Fields.Item("shipped").Value = "Cancelled" OR  rsAssignInvoices.Fields.Item("shipped").Value = "ON HOLD" then %>
                            <span class="font-weight-bold text-danger ml-2"><%= rsAssignInvoices.Fields.Item("shipped").Value %></span>
                            <% end if %>
                        
                        </label>
                      </div>
        </td>
    </tr>
<%
rsAssignInvoices.movenext()
wend
%>
</table>

<% else %>
<div class="alert alert-info mt-3">You are not logged in</div>

<%
end if ' user must be logged in to select invoices %>
</body>
</html>