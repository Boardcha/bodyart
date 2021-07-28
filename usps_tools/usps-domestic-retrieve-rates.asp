<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="usps_connection.asp" -->

<%
' https://www.usps.com/business/web-tools-apis/rate-calculator-api.htm

Set usps_xmlhttp = Server.CreateObject("MSXML2.ServerXMLHTTP") 
strReq = "<RateV4Request USERID=""" & usps_username & """>" _
& "<Package ID=""1"">" _
    & "<Service>First Class</Service>" _
    & "<FirstClassMailType>FLAT</FirstClassMailType>" _
    & "<ZipOrigination>78626</ZipOrigination>" _
    & "<ZipDestination>20770</ZipDestination>" _
    & "<Pounds>0</Pounds>" _
    & "<Ounces>8</Ounces>" _
    & "<Container>VARIABLE</Container>" _
    & "<Size>REGULAR</Size>" _
    & "<Machinable>true</Machinable>" _
& "</Package>" _
& "<Package ID=""2"">" _
    & "<Service>Priority Commercial</Service>" _
    & "<ZipOrigination>78626</ZipOrigination>" _
    & "<ZipDestination>20770</ZipDestination>" _
    & "<Pounds>0</Pounds>" _
    & "<Ounces>16</Ounces>" _
    & "<Container>SM FLAT RATE BOX</Container>" _
    & "<Size>REGULAR</Size>" _
    & "<Machinable>true</Machinable>" _
& "</Package>" _
& "<Package ID=""3"">" _
        & "<Service>Priority Mail Express</Service>" _
        & "<ZipOrigination>78626</ZipOrigination>" _
        & "<ZipDestination>20770</ZipDestination>" _
        & "<Pounds>0</Pounds>" _
        & "<Ounces>16</Ounces>" _
        & "<Container>FLAT RATE ENVELOPE</Container>" _
        & "<Size>REGULAR</Size>" _
        & "<Machinable>true</Machinable>" _
        & "</Package>" _
& "</RateV4Request>"


usps_xmlhttp.Open "POST","https://secure.shippingapis.com/ShippingAPI.dll?API=RateV4&XML=" & strReq, false
usps_xmlhttp.send

		usps_response = usps_xmlhttp.responseText 

	Set mydoc= Server.CreateObject("Microsoft.xmlDOM") 
        mydoc.loadxml( usps_response )
        
    'response.write usps_response & "<br><br>"

    Set rates_nodelist = mydoc.documentElement.selectNodes("Package") 

i = 0
for each node in rates_nodelist
    'Response.Write "<br>" & node.nodeName & "  =  " & node.text & "<br />" & vbCrLf

    If not(rates_nodelist.Item(i).selectSingleNode("@ID") is nothing) then
        if rates_nodelist.Item(i).selectSingleNode("@ID").Text = 1 then
    %>
    <br/>First class mail 
    <% end if
    if rates_nodelist.Item(i).selectSingleNode("@ID").Text = 2 then
    %>
    <br/>Priority mail 
    <% end if
    if rates_nodelist.Item(i).selectSingleNode("@ID").Text = 3 then
    %>
    <br/>Express mail 
    <% end if
    end if  

    If not(rates_nodelist.Item(i).selectSingleNode("Postage/CommercialRate") is nothing) then
    %>
    ---- Commercial rate: $<%= rates_nodelist.Item(i).selectSingleNode("Postage/CommercialRate").Text %>
    <%
    else
    ' Display retail rates if there's no commercial rate
        If not(rates_nodelist.Item(i).selectSingleNode("Postage/Rate") is nothing) then
        %>
        ---- Retail Rate: $<%= rates_nodelist.Item(i).selectSingleNode("Postage/Rate").Text %>
        <%
        end if  
    end if

i = i + 1
next
%>


