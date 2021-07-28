<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
'==== PAGE HAS BEEN BOOSTRAPPED =======
bootstrapped = "yes"

Set objCmd = Server.CreateObject ("ADODB.Command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT * FROM dbo.TBL_PurchaseOrders where po_hide = 0 AND Received = 'N' ORDER BY Received ASC, PurchaseOrderID DESC" 
objCmd.Parameters.Append(objCmd.CreateParameter("string_id",3,1,12,request("invoiceid")))
Set rsGetPurchaseOrders = objCmd.Execute()
%>
<html>
<head>
<title>Update Label Queries</title>
</head>
<body>
<!--#include file="admin_header.asp"-->

<div class="p-3">
  <div class="card mt-3">
    <div class="card-header h5">
      Alter label queries
    </div>
    <div class="card-body">
      <!--#include file="labels/inc-update-label-queries.asp" -->
    </div>
  </div> 

<table class="table table-striped table-borderless table-hover mt-5">
        <thead class="thead-dark">
                <tr>
                  <th>Date</th>
                  <th>Brand</th>
                  <th>Update Labels</th>
                </tr>
              </thead>
              <tbody>
<% 
While NOT rsGetPurchaseOrders.EOF
%>
         <tr>     
             <td>
    <%= FormatDateTime(rsGetPurchaseOrders.Fields.Item("DateOrdered").Value,2)%>
</td> 
<td>
    <strong><%=(rsGetPurchaseOrders.Fields.Item("Brand").Value)%></strong>
</td>
<td>
    <a href="barcodes_modifyviews.asp?ID=<%=(rsGetPurchaseOrders.Fields.Item("PurchaseOrderID").Value)%>&type=new_po_system">Update Labels</a>
</td>
</tr> 
                <% 
  rsGetPurchaseOrders.MoveNext()
Wend
%>
              </tbody>
</table>
</div>