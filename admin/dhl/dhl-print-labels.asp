<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/Connections/bodyartforms_sql_ADMIN.asp" -->
<style>
 /* style sheet for "letter" printing */
    @page {
        size: 4in 6in;
        margin: 0
    }
</style>
<%
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = MM_bodyartforms_sql_STRING

if request.querystring("request_amount") = "all" then

    objCmd.CommandText = "SELECT ID, dhl_package_id, dhl_base64_shipping_label FROM sent_items WHERE ship_code = N'paid' AND shipped = N'Pending shipment' AND dhl_base64_shipping_label <> '' AND (shipping_type LIKE '%DHL%') ORDER BY CASE WHEN PackagedBy = '' THEN 'aa' ELSE PackagedBy END, CASE WHEN shipping_type LIKE '%office%' THEN 1 WHEN autoclave = 1 THEN 2 WHEN shipping_type LIKE '%express%' THEN 4 WHEN shipping_type LIKE '%ups%' THEN 5 WHEN shipping_type LIKE '%USPS Priority mail%' THEN 6 WHEN shipping_type LIKE '%max%' THEN 7 WHEN  (shipping_type = 'DHL Basic mail') THEN 8 WHEN  (shipping_type = 'USPS First Class Mail') THEN 9 WHEN  (shipping_type = 'DHL GlobalMail Packet Priority') THEN 10 WHEN  (shipping_type = 'DHL GlobalMail Parcel Priority') THEN 11 WHEN  (shipping_type LIKE '%global basic%') THEN 12 ELSE 20 END DESC, ID DESC"

end if

if request.querystring("request_amount") = "packer" then

    objCmd.CommandText = "SELECT ID, dhl_package_id, dhl_base64_shipping_label FROM sent_items WHERE PackagedBy = ? AND ship_code = N'paid' AND shipped = N'Pending shipment' AND dhl_base64_shipping_label <> '' AND (shipping_type LIKE '%DHL%') ORDER BY CASE WHEN PackagedBy = '' THEN 'aa' ELSE PackagedBy END, CASE WHEN shipping_type LIKE '%office%' THEN 1 WHEN autoclave = 1 THEN 2 WHEN shipping_type LIKE '%express%' THEN 4 WHEN shipping_type LIKE '%ups%' THEN 5 WHEN shipping_type LIKE '%USPS Priority mail%' THEN 6 WHEN shipping_type LIKE '%max%' THEN 7 WHEN  (shipping_type = 'DHL Basic mail') THEN 8 WHEN  (shipping_type = 'USPS First Class Mail') THEN 9 WHEN  (shipping_type = 'DHL GlobalMail Packet Priority') THEN 10 WHEN  (shipping_type = 'DHL GlobalMail Parcel Priority') THEN 11 WHEN  (shipping_type LIKE '%global basic%') THEN 12 ELSE 20 END DESC, ID DESC"
    objCmd.Parameters.Append(objCmd.CreateParameter("packer",200,1,30, request.querystring("packer") ))

end if


if request.querystring("request_amount") = "single" then

    objCmd.CommandText = "SELECT dhl_package_id, dhl_base64_shipping_label FROM sent_items WHERE ID = ?"
    objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,12,request.querystring("invoiceid")))

end if

Set rsGetOrders = objCmd.Execute()

While NOT rsGetOrders.EOF
%>

<img src="data:image/png;base64, <%= rsGetOrders.Fields.Item("dhl_base64_shipping_label").Value %>" style="width:4in;height:6in">
<div style="page-break-after: always"></div>

<%
rsGetOrders.MoveNext()
Wend
%>
