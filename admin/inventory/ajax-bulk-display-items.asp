<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
'===== GENERATE INTEGER TO TEMP STORE WITH ADDED DATA UNTIL THE ORDER IS FINALIZED AND THEN CREATE WITH A PERMANENT PURCHASE ORDER NUMBER ===========================================
if request.Cookies("bulk-po-id") = "" then
    Response.Cookies("bulk-po-id") = Day(Now()) & Month(Now()) & Hour(Now()) & Minute(Now())
    Response.Cookies("bulk-po-id").Expires = DATE + 7
end if

    set objCmd = Server.CreateObject("ADODB.command")
    objCmd.ActiveConnection = DataConn
    objCmd.CommandText = "SELECT jewelry.ProductID, ProductDetailID, gauge, length, ProductDetail1, title, picture, qty FROM ProductDetails INNER JOIN jewelry ON ProductDetails.ProductID = jewelry.ProductID WHERE jewelry.ProductID = ?"
    objCmd.Parameters.Append(objCmd.CreateParameter("id",3,1,10, request("productid") ))
    Set rsGetVariants = objCmd.Execute()
%>
<div class="my-3">
<img class="img-fluid pull-left mr-3" src="https://bodyartforms-products.bodyartforms.com/<%= rsGetVariants("picture") %>"><h5 class="clearfix"><%= rsGetVariants("title") %></h5>
</div>

<table  class="table table-sm table-striped table-hover mt-2">
	<thead class="thead-dark">  
	<tr>
        <th>Quantity to pull</th>
        <th>Currently on hand</th>
        <th>Variant information</th>
    </tr>
</thead>
<%
While NOT rsGetVariants.EOF
%>
	
        <tr>
            <td class="form-inline">
                <input class="form-control form-control-sm mr-3" type="text" id="qty_<%= rsGetVariants("ProductDetailID") %>"><button class="btn btn-sm btn-secondary btn-add-item" type="button" data-id="<%= rsGetVariants("ProductDetailID") %>" data-on-hand-qty="<%= rsGetVariants("qty") %>">Add to pull list</button><span class="ml-2 msg-btn-add-<%= rsGetVariants("ProductDetailID") %>"></span>
            </td>
            <td width="15%"><span class="badge badge-success font-weight-bold p-2" id="on-hand-<%= rsGetVariants("ProductDetailID") %>"><%= rsGetVariants("qty") %></span></td>
            <td width="25%"><%= rsGetVariants("gauge") %>&nbsp;<%= rsGetVariants("length") %>&nbsp;<%= rsGetVariants("ProductDetail1") %></td>
        </tr>
<%
rsGetVariants.movenext()
wend
%>
</table>
<%
Set rsGetUser = nothing
DataConn.Close()
%>