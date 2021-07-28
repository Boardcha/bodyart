<!--#include virtual="/template/inc_includes_ajax.asp" -->
<%
if request("gauge") <> "" and request("gauge") <> "All" then
	sql_gauge = "AND ProductDetails.Gauge = ?"
	filter_gauge = request("gauge")
	filter_gauge_link = "&amp;gauge=" & server.URLEncode(request("gauge"))
end if

if request("color") <> "" and request("color") <> "All" then
	sql_color = "AND ProductDetails.ProductDetail1 = ?"
	filter_color = request("color")
	filter_color_link = "&amp;color=" & server.URLEncode(request("color"))
end if

	' Get all the photos for the product to display
	Set objCmd = Server.CreateObject ("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT TBL_PhotoGallery.PhotoID, TBL_PhotoGallery.ProductID, TBL_PhotoGallery.DetailID, TBL_PhotoGallery.filename, TBL_PhotoGallery.thumb_filename, TBL_PhotoGallery.description, ProductDetails.Gauge, ProductDetails.Length, ProductDetails.ProductDetail1, TBL_PhotoGallery.status, jewelry.title, TBL_PhotoGallery.DateSubmitted FROM TBL_PhotoGallery INNER JOIN ProductDetails ON TBL_PhotoGallery.DetailID = ProductDetails.ProductDetailID INNER JOIN jewelry ON TBL_PhotoGallery.ProductID = jewelry.ProductID WHERE status = 1 AND TBL_PhotoGallery.ProductID = ? " & sql_gauge & " " & sql_color & " ORDER BY PhotoID DESC"
	objCmd.Parameters.Append objCmd.CreateParameter("param1", 3, 1, 10, request("ProductID")) ' adDouble
	if request("gauge") <> "" and request("gauge") <> "All" then
		objCmd.Parameters.Append(objCmd.CreateParameter("filter_gauge",200,1,15,filter_gauge))
	end if
	if request("color") <> "" and request("color") <> "All" then
		objCmd.Parameters.Append(objCmd.CreateParameter("filter_color",200,1,75,filter_color))
	end if  

    set rsGetPhotos = Server.CreateObject("ADODB.Recordset")
    rsGetPhotos.CursorLocation = 3 'adUseClient
    rsGetPhotos.Open objCmd
    rsGetPhotos.PageSize = 20 ' not using (possibly needed for pagination)
    intPageCount = rsGetPhotos.PageCount ' not using (possibly needed for pagination)

    Select Case Request("Action")
	case "<<"
		intpage = 1
	case "<"
		intpage = Request("intpage")-1
		if intpage < 1 then intpage = 1
	case ">"
		intpage = Request("intpage")+1
		if intpage > intPageCount then intpage = IntPageCount
	Case ">>"
		intpage = intPageCount
	case else
		intpage = 1
end select


 If Not rsGetPhotos.EOF Then %>	
 <!--#include file="inc-gallery-paging.asp" -->
 <div class="baf-carousel" id="customer-photos">
	<!--#include virtual="/gallery/inc-photo-date-switch-links.asp" -->
			<% 
			rsGetPhotos.AbsolutePage = intPage '======== PAGING
For intRecord = 1 To rsGetPhotos.PageSize
			%>
				<a href="https://<%= DomainLink %>/<%= rsGetPhotos.Fields.Item("filename").Value %>"  data-fancybox="customer-photos" data-caption="<%= Server.HTMLEncode(rsGetPhotos.Fields.Item("Gauge").Value & " " & rsGetPhotos.Fields.Item("Length").Value & " " & rsGetPhotos.Fields.Item("ProductDetail1").Value) & " -- Photo # " & rsGetPhotos.Fields.Item("PhotoID").Value %>"><img class="w-auto mr-1 lazyload" style="height: 150px" src="https://<%= DomainLink %>/thumb_<%= rsGetPhotos.Fields.Item("filename").Value %>" alt="Customer photo" title="<%= Server.HTMLEncode(rsGetPhotos.Fields.Item("Gauge").Value & " " & rsGetPhotos.Fields.Item("Length").Value & " " & rsGetPhotos.Fields.Item("ProductDetail1").Value) & " -- Photo # " & rsGetPhotos.Fields.Item("PhotoID").Value %>"/></a>
			<% 
			rsGetPhotos.MoveNext()
If rsGetPhotos.EOF Then Exit For  ' ====== PAGING
Next ' ====== PAGING
			%>
			</div><!-- photo carousel -->
			
			<% else %>
<div class="alert alert-danger">NO PHOTOS FOUND matching &nbsp;&nbsp;&nbsp;<%= filter_gauge %>&nbsp;&nbsp;<%= filter_color %></div>
    <% End If ' Not rsGetPhotos.EOF 
    
    
    DataConn.Close()
    Set DataConn = Nothing
        %>