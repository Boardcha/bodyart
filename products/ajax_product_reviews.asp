<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<!--#include virtual="/template/inc_includes_ajax.asp" -->
<% 
sql_gauge = ""
filter_gauge = ""
filter_gauge_link = ""
filter_color = ""
filter_color_link = ""
sql_rating = ""
filter_rating = ""
filter_rating_link = ""
	
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

if request("filter_rating") <> "" then
	sql_rating = "AND review_rating = ?"
	filter_rating = request("filter_rating")
	filter_rating_display = "&nbsp;&nbsp;&nbsp;" & request("filter_rating") & " stars"
	filter_rating_link = "&amp;filter_rating=" & server.URLEncode(request("filter_rating"))
end if

' Get the top 5 newest reviews to display at bottom
Set objCmd = Server.CreateObject ("ADODB.Command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT TBLReviews.ProductID, TBLReviews.ReviewID, TBLReviews.review, TBLReviews.review_rating, TBLReviews.status, TBLReviews.name, TBLReviews.date_posted, TBLReviews.comments, ProductDetails.Gauge, ProductDetails.Length, ProductDetails.ProductDetail1 FROM TBLReviews LEFT OUTER JOIN ProductDetails ON TBLReviews.DetailID = ProductDetails.ProductDetailID WHERE review <> '' AND  TBLReviews.ProductID = ? " & sql_gauge & " " & sql_color & " " & sql_rating & " AND (TBLReviews.status = N'accepted' or TBLReviews.status IS NULL) ORDER BY ReviewID DESC" 
objCmd.Parameters.Append objCmd.CreateParameter("productid", 3, 1, 12, request("ProductID"))
if request("gauge") <> "" and request("gauge") <> "All" then
	objCmd.Parameters.Append(objCmd.CreateParameter("filter_gauge",200,1,15,filter_gauge))
end if
if request("color") <> "" and request("color") <> "All" then
	objCmd.Parameters.Append(objCmd.CreateParameter("filter_color",200,1,75,filter_color))
end if
if request("filter_rating") <> "" then
	objCmd.Parameters.Append(objCmd.CreateParameter("filter_rating",3,1,12,filter_rating))
end if

set rsGetReview = Server.CreateObject("ADODB.Recordset")
rsGetReview.CursorLocation = 3 'adUseClient
rsGetReview.Open objCmd
rsGetReview.PageSize = 5 ' not using (possibly needed for pagination)
intPageCount = rsGetReview.PageCount ' not using (possibly needed for pagination)

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


if NOT rsGetReview.EOF then 
%>
<div class="text-center">
<!--#include file="inc_reviews_paging.asp" -->
</div>
<%

rsGetReview.AbsolutePage = intPage '======== PAGING
For intRecord = 1 To rsGetReview.PageSize
%>
<div class="my-3">
<div class="bg-light p-2 rounded-top border">
	<% l = 0
			do until l = rsGetReview.Fields.Item("review_rating").Value %>
			   <i class="fa fa-star text-warning"></i>
		   <% l = l + 1 
		   loop %>
		   <span class="small text-secondary ml-1">
				<%= rsGetReview.Fields.Item("review_rating").Value %>
		</span>
	<span class="ml-2 text-secondary small">
	<span class="mr-4">
			<% if rsGetReview.Fields.Item("date_posted").Value <> "" then %>Posted on <%= rsGetReview.Fields.Item("date_posted").Value %><% end if %></span>
			<% if rsGetReview.Fields.Item("name").Value <> "" then %>&nbsp;&nbsp; by 
				<%= rsGetReview.Fields.Item("name").Value %>
			<% end if %>
	</span>	
</div>
<div class="p-2 rounded-bottom border-bottom border-left border-right">
		<% if rsGetReview.Fields.Item("review").Value <> "" then %>
		<span class="badge badge-warning font-weight-normal small mr-2"><i class="fa fa-verified mr-1"></i>Verified purchase</span>
		<span class="badge badge-secondary">
			<%= rsGetReview.Fields.Item("Gauge").Value %>&nbsp;<%=(rsGetReview.Fields.Item("Length").Value)%>&nbsp;<%=(rsGetReview.Fields.Item("ProductDetail1").Value)%>
		</span>
		<div class="mt-2">
				<%= rsGetReview.Fields.Item("review").Value %>
			</div>
				
		
				<% If (rsGetReview.Fields.Item("comments").Value) <> "" then %>
				<div class="alert alert-info"><span class="font-weight-bold">BAF comment: </span><%=(rsGetReview.Fields.Item("comments").Value)%></div>      
			  <% end if %>
			<div class="review-helpful d-none">
			Was this review helpful to you?<i class="fa fa-thumbs-up fa-lg"></i><i class="fa fa-thumbs-down fa-lg"></i>
			</div>
			 
		<% else %>
			<span class="review_noreview">No text review posted</span>
		<% end if %>
</div>
</div><!-- review margins -->
    <% 
rsGetReview.MoveNext()
If rsGetReview.EOF Then Exit For  ' ====== PAGING
Next ' ====== PAGING
%>
<!--#include file="inc_reviews_paging.asp" -->
<% else %>
<div class="alert alert-danger">NO RATINGS FOUND matching &nbsp;&nbsp;&nbsp;<%= filter_gauge %>&nbsp;&nbsp;<%= filter_color %>&nbsp;&nbsp;<%= filter_rating_display %></div>
<%
end if ' if recordset is not empty	


DataConn.Close()
Set DataConn = Nothing
%>