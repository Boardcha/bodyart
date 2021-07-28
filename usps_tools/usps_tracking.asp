<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="usps_connection.asp" -->

<%
	var_id = request.querystring("id")
%>
<div class="font-weight-bold pb-2">
	Tracking # <%= var_id %>
</div>
<%

	Set usps_xmlhttp = Server.CreateObject("MSXML2.ServerXMLHTTP") 
	usps_xmlhttp.Open "POST","http://production.shippingapis.com/ShippingAPI.dll?API=TrackV2&XML=<TrackRequest USERID=""" & usps_username & """><TrackID ID=""" & var_id & """></TrackID></TrackRequest>", false
	usps_xmlhttp.send

		usps_response = usps_xmlhttp.responseText 

	Set mydoc= Server.CreateObject("Microsoft.xmlDOM") 
		mydoc.loadxml( usps_response )

	Set pkg_nodelist = mydoc.documentElement.selectNodes("TrackInfo") 


'	for each node in pkg_nodelist
'	Response.Write node.nodeName & "  =  " & node.text & "<br />" & vbCrLf
'	For Each att in node.Attributes
'	  Response.Write att.Name & "  =  " & att.text & "<br />" & vbCrLf
'	Next
' next

	If not(pkg_nodelist.Item(0).selectSingleNode("TrackSummary") is nothing) then
		track_summary = pkg_nodelist.Item(0).selectSingleNode("TrackSummary").Text
	else
		track_summary = "USPS has not updated the status of your package yet. Please give 24 hours for their system to update and check again later."
	end if
	%>
<div class="alert alert-info">
		<%= track_summary %>
</div>
	

<%		
		Set objLst = mydoc.getElementsByTagName("TrackDetail")
		SizeofObject = objLst.length-1

			For each elem in objLst
				set childNodes = elem.childNodes
				for each node in childNodes
%>
<div class="my-2">
	<%= node.text %>
</div>	
<%					
				next
			Next
%>
