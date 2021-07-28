<%@LANGUAGE="VBSCRIPT" %>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<!--#include virtual="Connections/authnet.asp"-->
<%
'==== PAGE HAS BEEN BOOSTRAPPED =======
bootstrapped = "yes"

Function pd(n, totalDigits) 
if totalDigits > len(n) then 
    pd = String(totalDigits-len(n),"0") & n 
else 
    pd = n 
end if 
End Function  

first_date = YEAR(Date()-10) & "-" & Pd(Month(date()-10),2) & "-" & Pd(DAY(date()-10),2) 
last_date = YEAR(Date()) & "-" & Pd(Month(date()),2) & "-" & Pd(DAY(date()),2) 

' Authorize.net get batches list
strReq = "<?xml version=""1.0"" encoding=""utf-8""?>" _
& "<getSettledBatchListRequest xmlns=""AnetApi/xml/v1/schema/AnetApiSchema.xsd"">" _
& MerchantAuthentication() _
& "<includeStatistics>true</includeStatistics>" _
& "<firstSettlementDate>" & first_date & "T23:00:00Z</firstSettlementDate>" _
& "<lastSettlementDate>" & last_date & "T23:00:00Z</lastSettlementDate>" _
& "</getSettledBatchListRequest>"

Set objGetBatches = SendApiRequest(strReq)
%>
<!DOCTYPE html> 
<html>
<head>
<link rel="stylesheet" type="text/css" href="../CSS/Admin.css" />
<title>Batches</title>
</head>
<body>
<!--#include file="admin_header.asp"-->
<div class="p-3">
        <h5 class="mb-3">Batch list</h5>

        <table class="table table-striped table-hover">
                <thead class="thead-dark">
                        <tr>
                          <th>Date</th>
                          <th>Type</th>
                          <th>Total (less refunds)</th>
                          <th>Refunds</th>
                        </tr>
                      </thead>
                      <tbody>
<%
If IsApiResponseSuccess(objGetBatches) Then
set objBatches = objGetBatches.SelectNodes("/*/api:batchList/api:batch")


For Each batch In objBatches
    'strDate = FormatDateTime(batch.selectSingleNode("api:settlementTimeLocal").Text,1)
    str = batch.selectSingleNode("api:settlementTimeLocal").Text
    x = Instr(str,"T")
        If x Then strDate = Left(batch.selectSingleNode("api:settlementTimeLocal").Text, x-1)
        strDate = FormatDateTime(DateAdd("d", -1, strDate),1)
    

    If not(batch.selectSingleNode("api:marketType") is nothing) then
        strType = "Credit cards"
        strBatchTotal = 0
        strRefundTotal = 0
    else
        strType = "PayPal"
        strBatchTotal = batch.selectSingleNode("api:statistics/api:statistic/api:chargeAmount").Text
        strRefundTotal = batch.selectSingleNode("api:statistics/api:statistic/api:refundAmount").Text
    end if

     
    strCardType = ""
    set objStatistics = batch.SelectNodes("api:statistics/api:statistic")
    For Each statistic In objStatistics
        If not(statistic.selectSingleNode("api:accountType") is nothing) then
                'strCardType = statistic.selectSingleNode("api:accountType").Text & " " & strCardType
                strBatchTotal = Ccur(statistic.selectSingleNode("api:chargeAmount").Text) + Ccur(strBatchTotal)
                strRefundTotal = Ccur(statistic.selectSingleNode("api:refundAmount").Text) + Ccur(strRefundTotal)
        end if
    next
    
    strBatchTotal = FormatNumber(strBatchTotal - strRefundTotal)

    
%>
                        <tr>     
                     <td><%= strDate %>
        </td> 
        <td>
                <%= strType %> <%= strCardType %>
        </td>
        <td>
          $<%= strBatchTotal %>
        </td>
                <td>
                $<%= strRefundTotal %>
              </td>
        </tr> 
<%
        set objStatistics = nothing
Next   
End If
%>       
                      </tbody>
        </table>

</div>
</body>
</html>
<%
DataConn.Close()
%>