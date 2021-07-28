<%@LANGUAGE="VBSCRIPT" %>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
'==== PAGE HAS BEEN BOOSTRAPPED =======
bootstrapped = "yes"

set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = MM_bodyartforms_sql_STRING
objCmd.CommandText = "SELECT ID, company, customer_first, customer_last, address, address2, city, state, zip, country, phone, email FROM sent_items WHERE (dhl_package_id IS NULL OR dhl_package_id = '') AND ship_code = N'paid' AND shipped = N'Pending shipment' AND giftcert_flag = 0 AND (shipping_type LIKE '%DHL%')"
Set rsGetManifests = objCmd.Execute()
%>
<!DOCTYPE html> 
<html>
<head>
<title>DHL Labels & Closeouts</title>
</head>
<body>
<!--#include file="admin_header.asp"-->
<div class="p-3">
        <h5>Current DHL Orders</h5>

        <table class="table table-sm table-striped table-borderless table-hover small">
            <thead class="thead-dark">
                <tr>
                    <th>Invoice</th>
                    <th>Name</th>
                    <th>Address</th>
                    <th>City</th>
                    <th>State</th>
                    <th>Zip</th>
                    <th>Country</th>
                    <th>Phone</th>
                    <th>Email</th>
                </tr>
            </thead>
              <% While NOT rsGetManifests.EOF %>
              <tr>
                  <td>
                      <a href="invoice.asp?ID=<%= rsGetManifests.Fields.Item("ID").Value %>" target="_blank"><%= rsGetManifests.Fields.Item("ID").Value %></a>
                    </td>
                <td>
                    <%= rsGetManifests.Fields.Item("customer_first").Value %>&nbsp;<%= rsGetManifests.Fields.Item("customer_last").Value %>
                </td>
                <td>
                    <%= rsGetManifests.Fields.Item("address").Value %>&nbsp;<%= rsGetManifests.Fields.Item("address2").Value %>
                </td>
                <td><%= rsGetManifests.Fields.Item("city").Value %></td>
                <td><%= rsGetManifests.Fields.Item("state").Value %></td>
                <td><%= rsGetManifests.Fields.Item("zip").Value %></td>
                <td><%= rsGetManifests.Fields.Item("country").Value %></td>
                <td><%= rsGetManifests.Fields.Item("phone").Value %></td>
                <td><%= rsGetManifests.Fields.Item("email").Value %></td>
                </tr>
              <%
              rsGetManifests.MoveNext()
              Wend
              %>
        </table>
</div>
</body>
</html>
<%
DataConn.Close()
%>