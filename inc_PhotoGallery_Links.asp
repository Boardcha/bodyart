<% if CDate(rsGetPhotos.Fields.Item("DateSubmitted").Value) < CDate("2/20/2011") then
DomainLink = "bodyartforms-gallery.bodyartforms.com" ' AMAZON S3
'DomainLink = "216.128.23.22/gallery/uploads"
Else
DomainLink = "www.bodyartforms.com/gallery/uploads"
End if
%>

<a rel="group" title="<%= Server.HTMLEncode(rsGetPhotos.Fields.Item("title").Value & " " & rsGetPhotos.Fields.Item("Gauge").Value & " " & rsGetPhotos.Fields.Item("Length").Value & " " & rsGetPhotos.Fields.Item("ProductDetail1").Value) %>" href="http://<%= DomainLink %>/<%= rsGetPhotos.Fields.Item("filename").Value %>" class="fancy customer_photo"><img src="http://<%= DomainLink %>/thumb_<%= Replace(rsGetPhotos.Fields.Item("filename").Value, ".JPG", ".jpg") %>" alt="<%= Server.HTMLEncode(rsGetPhotos.Fields.Item("title").Value & " " & rsGetPhotos.Fields.Item("Gauge").Value & " " & rsGetPhotos.Fields.Item("Length").Value & " " & rsGetPhotos.Fields.Item("ProductDetail1").Value) & " -- Photo # " & rsGetPhotos.Fields.Item("PhotoID").Value %>" title="<%= Server.HTMLEncode(rsGetPhotos.Fields.Item("title").Value & " " & rsGetPhotos.Fields.Item("Gauge").Value & " " & rsGetPhotos.Fields.Item("Length").Value & " " & rsGetPhotos.Fields.Item("ProductDetail1").Value) & " -- Photo # " & rsGetPhotos.Fields.Item("PhotoID").Value %>" class="ProductInfo_Thumbnail_Gallery" /></a>