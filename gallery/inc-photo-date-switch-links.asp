<% if CDate(rsGetPhotos.Fields.Item("DateSubmitted").Value) < CDate("2/20/2011") then
DomainLink = "bodyartforms-gallery.bodyartforms.com" ' AMAZON S3
'DomainLink = "216.128.23.22/gallery/uploads"
Else
DomainLink = "www.bodyartforms.com/gallery/uploads"
End if

DomainLink = "bodyartforms-gallery.bodyartforms.com"
%>