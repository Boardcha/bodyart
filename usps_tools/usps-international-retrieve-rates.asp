<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="usps_connection.asp" -->

<%
' https://www.usps.com/business/web-tools-apis/rate-calculator-api.htm

Set usps_xmlhttp = Server.CreateObject("MSXML2.ServerXMLHTTP") 
strReq = "<IntlRateV2Request USERID=""" & usps_username & """>" _
& "<Revision>2</Revision>" _   
& "<Package ID=""1"">" _
    & "<Pounds>0</Pounds>" _
    & "<Ounces>8</Ounces>" _
    & "<Machinable>true</Machinable>" _
    & "<MailType>FLATRATE</MailType>" _
    & "<ValueOfContents>10.00</ValueOfContents>" _
    & "<Country>Australia</Country>" _
    & "<Container>RECTANGULAR</Container>" _
    & "<Size>REGULAR</Size>" _
    & "<Width>7</Width>" _
    & "<Length>7</Length>" _
    & "<Height>2</Height>" _
    & "<Girth>2</Girth>" _
    & "<OriginZip>78626</OriginZip>" _
& "</Package>" _
& "</IntlRateV2Request>"


usps_xmlhttp.Open "POST","https://secure.shippingapis.com/ShippingAPI.dll?API=IntlRateV2&XML=" & strReq, false
usps_xmlhttp.send

		usps_response = usps_xmlhttp.responseText 

	Set mydoc= Server.CreateObject("Microsoft.xmlDOM") 
        mydoc.loadxml( usps_response )
        
    'response.write usps_response & "<br><br>"

    Set rates_nodelist = mydoc.documentElement.selectNodes("Package") 

i = 0
for each node in rates_nodelist
    Response.Write "<br>" & node.nodeName & "  =  " & node.text & "<br />" & vbCrLf

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

    If not(rates_nodelist.Item(i).selectSingleNode("Service/Postage") is nothing) then
    %>
    ---- Retail rate: $<%= rates_nodelist.Item(i).selectSingleNode("Service/Postage").Text %>
    <%
    end if

i = i + 1
next
%>


